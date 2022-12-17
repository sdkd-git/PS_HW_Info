

#Start Winrm Service
net start winrm
$hostnm = hostname
$computers = "127.0.0.1"
#Time for script run
$time = (Get-Date).ToString("yyyy-MM-dd")
#Path to export csv
$path = "\\10.0.10.21\DomainExternal\IT\Script\InventoryData\"

#Hostname Confirmation
write-host "$hostnm"
$reply = Read-Host -Prompt "Continue with the hostname ?[y/n]"
if ( $reply -eq 'n' ) { 
    $newhostname = Read-Host "Enter the new hostname"
    $dcusername = Read-Host "Enter Domain Username Domain\username"
    Rename-Computer -NewName $newhostname -Force -DomainCredential $dcusername
}

#For monitor Decode
function Decode {
  If ($args[0] -is [System.Array]) {
      [System.Text.Encoding]::ASCII.GetString($args[0])
  }
  Else {
      "Not Found"
  }
}

$Result = [Collections.ArrayList]@() 

 
  $liveComputers = [Collections.ArrayList]@()

  foreach ($computer in $computers) {
    if (Test-Connection -ComputerName $computer -Quiet -count 1) {
      $null = $liveComputers.Add($computer)
    }
    else {
      Write-Verbose -Message ('{0} is unreachable' -f $computer) -Verbose
    }
  }

  $liveComputers | ForEach-Object {
#Bios Info
    $biosProps = 'SerialNumber', 'ReleaseDate', 'SMBIOSBIOSVersion', 'SMBIOSMajorVersion',
    'SMBIOSMinorVersion'
    $bios1 = Get-CimInstance -Class Win32_Bios -ComputerName $_ -Property $biosProps 
#Monitors Info
    $Monitors = Get-WmiObject -Class WmiMonitorID -Namespace root\wmi #| Select-Object Manufacturername  #-ComputerName $_
      ForEach ($Monitor in $Monitors) {
        $monManufacturer = Decode $Monitor.ManufacturerName -notmatch 0
        $monName = Decode $Monitor.UserFriendlyName -notmatch 0
        $monSerial = Decode $Monitor.SerialNumberID -notmatch 0
     }
#Memory
$hardware = Get-CimInstance -Class Win32_ComputerSystem -ComputerName $_  
#$totalMemory = [math]::round($hardware.TotalPhysicalMemory/1GB, 2)
 $totalMemory = Get-CimInstance CIM_PhysicalMemory -ComputerName $_ | Measure-Object -Property capacity -Sum | %{[math]::round(($_.sum / 1GB),2)}
#OS info
    $osProps = 'SerialNumber', 'Caption', 'Version','InstallDate','OSArchitecture', 'BuildNumber'
    $os1 = Get-CimInstance -ComputerName $_ -ClassName Win32_OperatingSystem -Property $osProps
#NIC
    $adapterProps = 'IPAddress', 'IPSubnet', 'MACAddress', 'DHCPEnabled', 'Description'
    $adapter = Get-CimInstance -Class Win32_NetworkAdapterConfiguration -ComputerName $_ | 
    Select-Object -Property $adapterProps |
    Where-Object { $_.IPAddress -ne $null }
# Printer Info
    #$printerprops = 'Name', 'PortName', 'Type'
    #$printerinfo = Get-Printer -Computername $computers  | Select-Object -Property $printerprops
    $printerinfo = Get-CimInstance -Class Win32_Printer -Computername $_
    # Get Hard Disk Info
    $HDDinfo = Get-PhysicalDisk -CimSession $_ | Select-Object FriendlyName, SerialNumber, MediaType, HealthStatus
    $HDDsize = Get-PhysicalDisk -CimSession $_ | Select-Object @{n=" ";e={[math]::Round($_.Size/1GB,2)}}
    $drivespaceProps = 'label', 'freespace', 'driveletter'  
    $driveSpace = Get-CimInstance -Class Win32_Volume -ComputerName $_ -Filter 'drivetype = 3' -Property $drivespaceProps | 
    Select-Object -Property driveletter, label, @{LABEL = 'GBfreespace'
    EXPRESSION                                        = { '{0:N2}' -f ($_.freespace/1GB) } 
    } |
    Where-Object { $_.driveletter -match 'C:' }
#CPU Info    
    $cpuProps = 'Name', 'NumberOfCores', 'Manufacturer', 'DeviceID'
    $cpu = Get-CimInstance -Class Win32_Processor -ComputerName $_ -Property $cpuProps 
  
    $hotfixProps = 'HotFixID', 'Description', 'InstalledOn'
    $hotFixes = Get-CimInstance -ClassName Win32_QuickFixEngineering -ComputerName $_ -Property $hotfixProps | 
    Select-Object -Property $hotfixProps | 
    Sort-Object InstalledOn -Descending
   
    # create new custom object to keep adding store information to it
    $Result += New-Object -TypeName PSCustomObject -Property @{
      ScriptRunDate           = Get-Date
      ComputerName             = $hostnm.ToUpper()
      Manufacturer             = $hardware.Manufacturer
      Model                    = $hardware.Model
      SystemType               = $hardware.SystemType
      ProductName              = $os1.Caption
      OSVersion                = $os1.version
      BuildNumber              = $os1.BuildNumber
      OSArchitecture           = $os1.OSArchitecture
      OSSerialNumber           = $os1.SerialNumber
      InstallDate              = $os1.InstallDate
      SerialNumber             = $bios1.SerialNumber
      BIOSReleaseDate          = $bios1.ReleaseDate
      BIOSVersion              = $bios1.SMBIOSBIOSVersion
      MonitorMfgName           = $monManufacturer
      MonitorModel             = $monName
      MonitorSrNo              = $monSerial
      IPAddress                = ($adapter.IPAddress -replace (",", "\n") | Out-String)
      SubnetMask               = ($adapter.IPSubnet -replace (",", "\n") | Out-String)
      MACAddress               = ($adapter.MACAddress -replace (",", "\n") | Out-String)
      DHCPEnabled              = ($adapter.DHCPEnabled -replace (",", "\n") | Out-String)
      AdapterInfo              = ($adapter.Description -replace (",", "\n") | Out-String)
      Domain                   = $hardware.Domain
      TotalMemoryGB            = $totalMemory
      CFreeSpaceGB             = $driveSpace.GBfreespace
      HDDName                  = $HDDinfo.FriendlyName | Out-String
      HDDSrNo                  = $HDDinfo.Serialnumber | Out-String
      HDDType                  = $HDDinfo.MediaType | Out-String
      HDDHealth                = $HDDinfo.HealthStatus | Out-String
      HDDsize                  = $HDDsize | Out-String
      CPUManufacturer          = $cpu.Manufacturer | Out-String
      CPU                      = $cpu.Name | Out-String
      CPUCores                 = $cpu.NumberOfCores | Out-String
      CPUDeviceID              = $cpu.DeviceID | Out-String
      HotFixID                 = ($hotFixes.HotFixID -replace (",", "\n") | Out-String)
      Description              = ($hotFixes.Description -replace (",", "\n") | Out-String)
      InstalledOn              = ($hotFixes.InstalledOn -replace (",", "\n") | Out-String)
      Username                 = $env:UserName
      PrinterName              = $printerinfo.Name | Out-String
      #PrinterType              = $printerinfo.Type
      #PrinterPortname          = $printerinfo.PortName
    }
  }

# Column ordering, re-order if you like 
  $colOrder = 'ScriptRunDate','Username','ComputerName', 'Manufacturer', 'Model', 'SystemType', 'ProductName', 'OSVersion', 'BuildNumber', 'OSArchitecture', 'OSSerialNumber',
  'SerialNumber', 'BIOSReleaseDate', 'BIOSVersion', 'MonitorMfgName', 'MonitorModel','MonitorSrNo', 'InstallDate', 'Domain', 'IPAddress', 'SubnetMask', 'MACAddress',
  'DHCPEnabled','AdapterInfo', 'TotalMemoryGB', 'CFreeSpaceGB', 'HDDName','HDDSrNo','HDDType','HDDsize', 'CPU', 'CPUCores', 
  'CPUManufacturer', 'CPUDeviceID', 'HotFixID', 'Description', 'InstalledOn', 'PrinterName'
  #, 'PrinterType', 'PrinterPortname'
#Return your results
$Result | Select-Object -Property $colOrder | Export-csv -Path $path$time.csv -NoTypeInformation -Append
