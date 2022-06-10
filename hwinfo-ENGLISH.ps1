# | For this script is part of more advanced system combined of several modules 
# | it usually takes predefined PCName variable from another module
# | But if PCName is not defined (it means that script run as standalone module) 
# | write local system hostname as PCName variable
IF ([string]::IsNUllOrEmpty($PCName)) {
$PCName = $(hostname)
}

# | Set html-file name to store PC info as system name 
# | And set to write file to the same location where script is stored, using PSScriptRoot command
$FilePath = "$PSScriptRoot\$PCName.html" 

IF (![System.IO.File]::Exists("$FilePath")) {
# | If html file named as hostname does NOT exist in specified directory
# | First write to html-file CSS styles
Add-Content -Path $FilePath -Value '<head>'
Add-Content -Path $FilePath -Value '<style>'
Add-Content -Path $FilePath -Value '.div-main {width: 100%; height: 800px; padding: 10px; margin: 10px;}'  
Add-Content -Path $FilePath -Value '.div-table {display: inline-block; width: auto; height: 780px; border-left: 1px solid black; padding: 10px; float: left; margin-top: 10px; margin-left: 10px; margin-bottom: 10px;}' 
Add-Content -Path $FilePath -Value '.div-table-sec {display: inline-block; width: auto; height: 780px; border-left: 1px solid black; padding: 10px; margin-top: 10px; margin-bottom: 10px;}'
Add-Content -Path $FilePath -Value '.div-table-row {display: table-row;}'  
Add-Content -Path $FilePath -Value '.div-table-cell-date {width: auto; padding: 10px 40px 10px; border-top: 1px solid black; border-right: 1px solid black; float: right; color: red;}'  
Add-Content -Path $FilePath -Value '.div-table-cell-pcname {width: auto; padding: 10px 40px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;}'  
Add-Content -Path $FilePath -Value '.div-table-cell-long {width: auto; min-width: 180px; max-width: 400px; padding: 10px 20px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;}'  
Add-Content -Path $FilePath -Value '.div-table-cell {width: 180px; padding: 10px 20px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;}'  
Add-Content -Path $FilePath -Value '</head></style>'  
}
  
# | Write to html file main positioning div
# | Later in comments, if not specified otherwise, it is supposed that we write info to html file.
Add-Content -Path $FilePath -Value '<div class="div-main">'
  
# | Write first main div-table
Add-Content -Path $FilePath -Value '<div class="div-table">'
  
# | Write hostname
Add-Content -Path $FilePath -Value '<div class="div-table-row">System Name</div>'
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell-pcname">'
$PCName | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</div>'

# | Write date and time
Add-Content -Path $FilePath -Value '<div class="div-table-cell-date">'
Get-Date -Format "dd/MM/yy HH:mm" | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</div></div>' 
  
# | Write Motherboard Model and Manufacturer to Collection 
$BaseBoardInfo = Get-CimInstance -Query "Select Manufacturer, Product from Win32_BaseBoard"

# | Write table header for Motherboard info
Add-Content -Path $FilePath -Value '<div class="div-table-row">Motherboard</div>'

# | Write info on Manufacturer
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
$BaseBoardInfo.Manufacturer | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</div>'

# | Write info on motherboard model
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
$BaseBoardInfo.Product | Add-Content -Path $FilePath
Add-Content -Path $FilePath -Value '</div>'

# | Write info on BIOS/UEFI to Collection
$BIOSInfo = Get-CimInstance Win32_Bios 

# | Write BIOS/UEFI info
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value 'BIOS/UEFI Ver.:' 
Add-Content -Path $FilePath -Value $BIOSInfo.SMBIOSBIOSVersion
Add-Content -Path $FilePath -Value '</div></div>'

# | Write CPU DeviceID, Socket, Name, Number of Cores and Threads to Collection 
# | Use NumberOfLogicalProcesors because ThreadCount only works in W10
$CPUInfo = Get-CImInstance -Query "Select DeviceID, SocketDesignation, Name, NumberOfCores, NumberOfLogicalProcessors from Win32_Processor"

# | Write table header for info on CPUs
Add-Content -Path $FilePath -Value '<div class="div-table-row">CPU(s)</div>'  

# | Use cycle for sometimes there are several CPU you know
# | for each CPU 
foreach ($i in $CPUInfo) {

# | Write CPU ID in System 
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value $i.DeviceID
Add-Content -Path $FilePath -Value '</div>'

# | Write CPU Name 
Add-Content -Path $FilePath -Value '<div class="div-table-cell-long">'
Add-Content -Path $FilePath -Value $i.Name
Add-Content -Path $FilePath -Value '</div></div>'

# | Write Socket
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value $i.SocketDesignation
Add-Content -Path $FilePath -Value '</div>'

# | Write Cores/Threads
Add-Content -Path $FilePath -Value '<div class="div-table-cell-long">'
Add-Content -Path $FilePath -Value 'Cores and Threads: '
Add-Content -Path $FilePath -Value $i.NumberOfCores
Add-Content -Path $FilePath -Value '/'
Add-Content -Path $FilePath -Value $i.NumberOfLogicalProcessors
Add-Content -Path $FilePath -Value '</div></div>'
}

# | Write table header for Info on Operating Memory
Add-Content -Path $FilePath -Value '<div class="div-table-row">RAM</div>'

# | Write info on MemorySlot, Capacity, Speed, and Size 
$MemoryInfo = Get-CImInstance -Query "Select DeviceLocator, Capacity, Speed from Win32_PhysicalMemory"

# | For each RAM module
foreach ($i in $MemoryInfo) {

# | Write Slot number
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value $i.DeviceLocator
Add-Content -Path $FilePath -Value '</div>'

# | Write Capacity (Convert to Mb in advance)
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value $([Math]::Round($i.Capacity/1Mb))
Add-Content -Path $FilePath -Value 'Mb'
Add-Content -Path $FilePath -Value '</div>'

# | Write Speed 
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value $i.Speed 
Add-Content -Path $FilePath -Value 'MHz'
Add-Content -Path $FilePath -Value '</div></div>'
}

# | Write table header for Storage info 
Add-Content -Path $FilePath -Value '<div class="div-table-row">Storage Device(s)</div>'

# | Write Name, SerialNUmber, Size and Status for every storage device that match 'Fixed Hard Disk Media' 
# | To filter out Removable Media such as USB Stick and so on
$StorageInfo = Get-CImInstance -Query "Select Caption, SerialNumber, Size, Status from Win32_DiskDrive WHERE MediaType LIKE 'Fixed Hard Disk Media'"

# | For each Storage Device
foreach ($i in $StorageInfo) {

# | Write Device Name
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value $i.Caption
Add-Content -Path $FilePath -Value '</div>'

# | Serial number
Add-Content -Path $FilePath -Value '<div class="div-table-cell-long">'
Add-Content -Path $FilePath -Value 's/n: '
Add-Content -Path $FilePath -Value $i.SerialNumber
Add-Content -Path $FilePath -Value '</div></div>'

# | Size 
# | In advance divide Size by 1000000000 instead of 1Gb to get correct value
# | Because PowerShell Gb = 1024Mb and Storage Device Manufacturers Gb = 1000Mb
# | Also sometimes it is necessary to round final value for clean result
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value $([Math]::Round($i.Size/1000000000))
Add-Content -Path $FilePath -Value 'Gb'
Add-Content -Path $FilePath -Value '</div>'

# | Disk Status
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value $i.Status
Add-Content -Path $FilePath -Value '</div></div>'
}

# | Write table header for GPUs info
Add-Content -Path $FilePath -Value '<div class="div-table-row">Video Adapters</div>'

# | Get data on GPUs from regedit as collection of arrays
$GPUsInfo = Get-ItemProperty -path "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4D36E968-E325-11CE-BFC1-08002BE10318}\????" 

# | Get DeviceID of GPUs in system
# | Select only phisically presented devices to get rid of virtual devices created for RDP sessions and so on
$GPUDeviceIDs = Get-CIMInstance -Query "SELECT PNPDeviceID from Win32_VideoController WHERE PNPDeviceID LIKE 'PCI%'"

# | for each DeviceID
foreach ($i in $GPUDeviceIDs.PNPDeviceID) {

# | From GPUsInfo collection select array in which MatchingDeviceID element match current DeviceID in cycle
# | And store it as CurrentGPU variable
$CurrentGPU = $GPUsInfo | where { $i -like "$($_.MatchingDeviceID)*" }

# | Store curent GPU name in variable GPUName
$GPUName = $CurrentGPU.DriverDesc

# | Check if GPU memory is properly defined 
# | And if it is less than 1 smth went wrong or most likely system has Intel integrated GPU
# | for in this case qwMemorySize is not stored in regedit
IF ($CurrentGPU.'HardwareInformation.qwMemorySize' -lt 1) {
$GPUMemory = "N/A or DVMT"
} ELSE {
# | Otherwise, if memory size properly aquired from regedit 
# | Convert it to Mb and write to GPUMemory variable
$GPUMemory = ($CurrentGPU.'HardwareInformation.qwMemorySize')/1Mb
}

# | Write GPU Name
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value $GPUName
Add-Content -Path $FilePath -Value '</div>'

# | Write GPU Memory size
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value $GPUMemory
Add-Content -Path $FilePath -Value '</div></div>'
}


# | Write table header for Network Adapters info 
Add-Content -Path $FilePath -Value '<div class="div-table-row">Network Adapters</div>'

# | Write to collection Name and MAC-Address of physical Network Adapters 
$PhysicalNetworkAdaptes = Get-CIMInstance -Query "SELECT Name, MACAddress FROM Win32_NetworkAdapter WHERE PhysicalAdapter='True'"

# | For each Network Adapter in PhysicalNetworkAdapters collection
foreach ($i in $PhysicalNetworkAdaptes) {

# | Write Network Adapter Name
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell-long">'
Add-Content -Path $FilePath $i.Name 
Add-Content -Path $FilePath -Value '</div>'

# | Write Network Adapter MAC-Address
Add-Content -Path $FilePath -Value '<div class="div-table-cell-long">'
Add-Content -Path $FilePath $i.MACAddress
Add-Content -Path $FilePath -Value '</div></div>'
}

# | Close first div table with major hardware info
Add-Content -Path $FilePath -Value '</div>'
  
# | Open second div table for software info
Add-Content -Path $FilePath -Value '<div class="div-table-sec">'
  
# | Write table header for Operating System info
Add-Content -Path $FilePath -Value '<div class="div-table-row">Operating System</div>'

# | Get info on Operting System Architecture, Name and Version
$OperatingSystemInfo = Get-CIMInstance -Query " SELECT OSArchitecture, Caption, Version FROM Win32_OperatingSystem"

# | Write Operating System Name
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value 'Name' 
Add-Content -Path $FilePath -Value '</div>'
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value $OperatingSystemInfo.Caption
Add-Content -Path $FilePath -Value '</div></div>'

# | # | Write Operating System Version
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value 'Version' 
Add-Content -Path $FilePath -Value '</div>'
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value $OperatingSystemInfo.Version
Add-Content -Path $FilePath -Value '</div></div>'

# | Architecture
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value 'Architecture' 
Add-Content -Path $FilePath -Value '</div>'
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value $OperatingSystemInfo.OSArchitecture
Add-Content -Path $FilePath -Value '</div></div>'

# | Activation Status 
# | Do not use SoftwareLicensingProduct for it is more time consuming and less informative in our case
# | Get windows activation status as array and split it 
# | to output only second element which is activation status itself
$WindowsActivationStatus = ((cscript.exe /nologo "C:\windows\system32\slmgr.vbs" /xpr) -Split ":")

Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value 'Activation' 
Add-Content -Path $FilePath -Value '</div>'
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value $WindowsActivationStatus[2] 
Add-Content -Path $FilePath -Value '</div></div>'


# | Get info on MS Office if any 
# | Write to collection Name and Install location of every software where Name match Office and Vendor is Microsoft Corporation
$MSOfficeInfo = Get-CimInstance -Query "SELECT InstallLocation, Name, Vendor FROM Win32_Product WHERE Name LIKE '%Office%' AND Vendor LIKE '%Microsoft Corporation%'"

# | For each element in MSOfficeInfo collection if Name contain Standard, Professional or 365
# | Write Name and Install path to MSOfficeName and MSOfficePath variables accordingly
foreach ($i in $MSOfficeInfo) {
IF ($i.Name -match "Standard" -xor $i.Name -match "Professional" -xor $i.Name -match "365")
{
$MSOfficeName = $i.Name 
$MSOfficePath = $i.InstallLocation
}
}

# | If MSOfficePath is not defined or empty write "N/A" to MSOfficeName and MSOfficePath
IF ([string]::IsNUllOrEmpty($MSOfficePath)) {
$MSOfficeName = "N/A"
$MSOfficePath = "N/A"

# | Otherwise find out MS Office Name and Activation status
} ELSE {

# | Search for OSPP.VBS file and write path to it to variable
$OsppVbsPath = Get-ChildItem -Path $MSOfficePath -Filter OSPP.VBS -Recurse | foreach-object {$_.FullName}

# | Feed OsppVbsPath to command which output MS Office activation status
# | Then feed command output to Select-String to filter out only string with license status
# | And use -replace operator to get from string only needed part
$MSOfficeActivationStatus = ((cscript $OsppVbsPath /dstatus) | Select-String -Pattern "LICENSE STATUS") -replace "LICENSE STATUS:", "" -replace "-", ""

# | Write info on MS Office 
Add-Content -Path $FilePath -Value '<div class="div-table-row">MS Office</div>'
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value $MSOfficeName
Add-Content -Path $FilePath -Value '</div>'
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value $MSOfficeActivationStatus
Add-Content -Path $FilePath -Value '</div></div>'
}

# | Get info on Network Connections Settings
# | Write in collection MAC-Address, Connection Name, Connection Status for every Physical Adapter in System
$NetAdaptersConfig = Get-CimInstance -Query "Select MACAddress, NetConnectionID, NetConnectionStatus from Win32_NetworkAdapter where PhysicalAdapter='True'"

# | For each NetAdapter in NetAdapterConfig collection
foreach ($i in $NetAdaptersConfig) {

# | If connection is active
IF ($i.NetConnectionStatus -eq 2) {

# | Write connection name and add (Active) to specify currently active connection
Add-Content -Path $FilePath -Value '<div class="div-table-row">'
Add-Content -Path $FilePath -Value $i.NetConnectionID 
Add-Content -Path $FilePath -Value "(Active)"
Add-Content -Path $FilePath -Value '</div>'

# | Write MAC-Address of active connection
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value 'MAC Address'
Add-Content -Path $FilePath -Value '</div>'
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value $i.MACAddress
Add-Content -Path $FilePath -Value '</div></div>'

# | Write to variable current mac-address 
$CurrentMACAddress = $i.MACAddress

# | Get active connection network parameters from Win32_NetworkAdapterConfiguration filtering out by CurrentMACAddress
$ActiveNetConnection = Get-CimInstance -Query "Select MACAddress, IPAddress, IPSubnet, DefaultIPGateway from Win32_NetworkAdapterConfiguration where MACAddress like '$CurrentMACAddress'"

# | Write IP-Address of active connection
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value 'IP Address'
Add-Content -Path $FilePath -Value '</div>'
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value $ActiveNetConnection.IPAddress
Add-Content -Path $FilePath -Value '</div></div>'

# | Write network mask of active connection
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value 'Subnet Mask'
Add-Content -Path $FilePath -Value '</div>'
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value $ActiveNetConnection.IPSubnet
Add-Content -Path $FilePath -Value '</div></div>'

# | Write default gateway of active connection
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value 'Gateway'
Add-Content -Path $FilePath -Value '</div>'
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value $ActiveNetConnection.DefaultIPGateway
Add-Content -Path $FilePath -Value '</div></div>'

}
# | Otherwise (NetConnectionStatus NOT equal 2) connection inactive
ELSE {

# | Write connection name and add (Inactive)
Add-Content -Path $FilePath -Value '<div class="div-table-row">'
Add-Content -Path $FilePath -Value $i.NetConnectionID 
Add-Content -Path $FilePath -Value "(Inactive)"
Add-Content -Path $FilePath -Value '</div>'

# | Write MAC-Address of inactive Connection
Add-Content -Path $FilePath -Value '<div class="div-table-row"><div class="div-table-cell">'
Add-Content -Path $FilePath -Value 'MAC Address'
Add-Content -Path $FilePath -Value '</div>'
Add-Content -Path $FilePath -Value '<div class="div-table-cell">'
Add-Content -Path $FilePath -Value $i.MACAddress
Add-Content -Path $FilePath -Value '</div></div>'
}
}

# | Close second div table 
Add-Content -Path $FilePath -Value '</div>'
  
# | # | Close main positioning div table 
Add-Content -Path $FilePath -Value '</div>'
