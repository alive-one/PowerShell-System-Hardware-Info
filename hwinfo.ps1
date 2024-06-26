# | Configuration
# | Path to write data (Default is script root directory)
$FilePath = "$PSScriptRoot" 

# | Set Data Format Output. For example $JSON = 1 to export Data in *.json format. Multiple Choice is Possible
$JSON = 1

# | CSV Export Settings
$CSV = 1

# | Set standard CSV delimiter which is comma (,) by default
$StringCSVDelimiter = ','
$DictionaryCSVDelimiter = '","'

# | Yet some programs, such as MS Excel, use regional settings when import CSV Data
# | If you prefer to use your regional settings for CSV export, UNCOMMENT two lines below
#$StringCSVDelimiter = (Get-Culture).TextInfo.ListSeparator
#$DictionaryCSVDelimiter = "`"$($StringCSVDelimiter)`""

# | XML Export Settings
$XML = 1

# | HTML(GUI) Export Settings
$HTML = 1

# | As hwinfo.ps1 is a part of more complex system, current script must already have $PCName defined from previuos module "Autoren"
# | If not then this script is launched as standalone module and need to define $PCName as local System Name
IF ([string]::IsNUllOrEmpty($PCName)) {
[string]$PCName = $(hostname)
}

# | Predefined RAM Codes from Win32_PhysicalMemory.MemoryType and SMBIOSMemoryType
$RAMTypes = @{0 = 'Unknown'; 1 = 'Other'; 20 = 'DDR'; 21 = 'DDR2'; 22 = 'DDR2 FB-DIMM'; 24 = 'DDR3'; 26 = 'DDR4'; 34 = 'DDR5'}

# | Ordered Dictionary to Store Data in desired Order
$PCInfo = [ordered]@{

# | Local System Date (Be wary for it may differ from Accurate Date sometimes)
'LocalDate' = [string]$(Get-Date -Format "dd/MM/yyyy HH:mm"); 

# | Literally. 'Nuff said 
'SystemName' = $PCName;
}

# | Use Queres where it is possbile because queries are fastest method to get CIM-Instance Data
# | Baseboard Manufacturer and Model
$BaseBoardQuery = Get-CimInstance -Query "Select Manufacturer, Product from win32_Baseboard"

# | Socket, Name, etc. CPU staff
$CPUQuery = Get-CimInstance -Query "Select DeviceID, Manufacturer, SocketDesignation, Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed from Win32_Processor"

# | Slot Name, Size, Speed, MemoryType
$RAMQuery = Get-CimInstance -Query "Select DeviceLocator, BankLabel, Capacity, Speed, MemoryType, SMBIOSMemoryType from Win32_PhysicalMemory"

# | RAM Type stored in MemoryType, SMBIOSMemoryTYpe or Both. So Use .where() method to Filter Out only first suitable value
[string]$RAMType = $($RAMTypes.[int]$RAMQuery.MemoryType[0], $RAMTypes.[int]$RAMQuery.SMBIOSMemoryType[0]).where({$_ -notlike 'Undefined'},'First')

# | All RAM Slots Available (No Matter Populaterd or Not)
$RAMSlotsCount = $(Get-CimInstance -Query "Select MemoryDevices from Win32_PhysicalMemoryArray").MemoryDevices

# | All RAM Available
[int]$TotalRAMCapacity = 0

# Calculate All RAM Installed in Gb
foreach ($RAMModuleSize in $RAMQuery.Capacity) {
$TotalRAMCapacity = $TotalRAMCapacity + $RAMModuleSize / 1Gb
}

# | Physical Storage Devices excluding USB, SD-Cards, other Removables
# | Use Get-PhysicalDisk because only this cmdlet return proper values for BusType and MediaType, i.e. (SATA) and SSD accordingly
$StorageQuery = Get-PhysicalDisk | where {$_.BusType -notlike "USB" -xor "SD" -xor "MMC" -xor "Unknown" -xor "Virtual"} | Select DeviceID, BusType, MediaType, FriendlyName, Size, SerialNumber | Sort-Object -Property DeviceID

# | GPU Name, DeviceID and Proper GPUMemory amount from Registry
$GPUsRegQuery = Get-ItemProperty -path "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4D36E968-E325-11CE-BFC1-08002BE10318}\????" | Select -Property MatchingDeviceID, DriverDesc, HardwareInformation.qwMemorySize 

# | DeviceIDs of all currently installed GPUs
# | Store GPU IDs Strings to Array to use Array Indexes for GPU devices enumeration later
$InstalledGPUs = @(Get-CIMInstance -Query "SELECT PNPDeviceID, AdapterCompatibility from Win32_VideoController WHERE PNPDeviceID LIKE 'PCI%'")

# | Query Physical Adapters Only and store in Array to use Array Indexes for NetAdapters enumeration
$PhysicalAdaptersQuery = @(Get-CimInstance -ClassName MSFT_NetAdapter -Namespace ROOT\StandardCimv2 | Where-Object {!$_.Virtual})

# | Query Network Connections with IPEnbled 
$NetConnectionsQuery = @(Get-CimInstance -Query "Select Description, IPAddress, IPSubnet, DefaultIPGateway, MACAddress from Win32_NetworkAdapterConfiguration WHERE IPEnabled = 1")

# | Add Info to PCInfo Dictionary
# | Baseboard 
$PCInfo["BaseBoard"] = [ordered]@{

'Manufacturer' = $BaseBoardQuery.Manufacturer; 

'Model' = $BaseBoardQuery.Product; 

'RAMSLots' = "$($RAMSlotsCount) RAM Slots";
}

# | CPUs (There can be a few, you know)
foreach ($CPU in $CPUQuery) {

# | IF CPU Manufacturer is Intel or AMD
IF ([string]$CPU.Manufacturer -match "Intel|AMD") {

# | Leave only Intel or AMD as Manufacturer String using predefined array Matches[] which contains results of last comparison
$CleanCPUManufacturer = [string]$Matches[0]

# | Remove Manufacturer from CPU Name with regexp to get rid of data doubling;
$CleanCPUName = [string]$CPU.Name -replace ".*?($CleanCPUManufacturer).*? ", ""

# | Just incase LongSoon, some ARM or even Elbrus, you know...
} ELSE {
# | Leave CPU Manufcaturer as is
[string]$CleanCPUManufacturer = $CPU.Manufacturer

# | Leave CPU Name as is
[string]$CleanCPUName = $CPU.Name
}

# | Use each CPU.DeviceID as name for Ordered Dictionary which contains CPU Name and Cores/Threads Count. And add this dictionary to PCInfo
$PCInfo[$CPU.DeviceID] = [ordered]@{

'Socket' = $CPU.SocketDesignation;

'Manufacturer' = $CleanCPUManufacturer;

'Name' = $CleanCPUName;

'Cores/Threads' = "$($CPU.NumberOfCores) Cores / $($CPU.NumberOfLogicalProcessors) Threads";
}
}

# | Common RAM Info 
$PCInfo['RAM'] = [ordered]@{

'RAMType' = $RAMType;

'TotalRAM' = "$($TotalRAMCapacity) Gb";

'Speed' = "$($RAMQuery[0].Speed) Mhz"
}


# | Individual RAM Modules Info
foreach ($RAMModule in $RAMQuery){

# | Remove spaces in Memory BANK names since names with spaces are not allowed as XML Tags
$PCInfo[("RAM$($RAMModule.BankLabel)").replace(" ", "")] = [ordered]@{

'ModuleSlot' = "$($RAMModule.DeviceLocator)";

'ModuleSize' = "$($RAMModule.Capacity / 1MB) Mb";
}
} 


# | Storage
foreach ($Device in $StorageQuery) {

# | From Win32_DiskDrive query PNPDeviceID, filtering by serial number of device currently in cycle
$DevicePNP = (Get-CimInstance -Query "Select PNPDeviceID, SerialNumber from Win32_DiskDrive" | where {$(($_.SerialNumber).Trim()) -like "*$($Device.SerialNumber)*"}).PNPDeviceID

# | Use Each Storage Device Unique Identifier as enumerator to name Ordered Dictionary which will contain all current Storage Device info
$PCInfo["Disk$($Device.DeviceID)"] = [ordered]@{

# | Define Bus and Storage Type
'(Bus)Type' = "($($Device.BusType))$($Device.MediaType)";

# | Use regex to filter out only manufacturer name from PNPDeviceID string
'Manufacturer' = [regex]::Match($DevicePNP, "VEN_(.+?)&").Groups[1].Value;

# | Use regex to remove redundant info such as Manufacturer name from Model name
'Model' = [regex]::Replace($Device.FriendlyName, $DeviceManufacturer, "", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase);

# | Use division by 1000000000 instead of Powershell 'Gb' to set 'proper' Disk Size
'Size' = [string]([Math]::Round($Device.Size/1000000000)) + " Gb";

}
}


# GPUs
foreach ($GPUDevice in $InstalledGPUs) {

# | Compare currently installed GPU Device IDs with corresponding values from Registry to get necessary registry branch with actual info on GPU Memory size
$CurrentGPU = $GPUsRegQuery | where {$GPUDevice.PNPDeviceID -like "$($_.MatchingDeviceID)*"}

# | Use ArrayIndex from InstalledGPUDeviceIDs Query to create unique Dictionary name to store GPUName and Memory Size.
$PCInfo["GPU$($InstalledGPUs.IndexOf($GPUDevice))"] = [ordered]@{

# | GPU Manufacturer
'Manufacturer' = $GPUDevice.AdapterCompatibility;

# | Remove Manufacturer Name from GPU Model Name since we have already store Manufacturer in appropriate separated field above
'Model' = [regex]::Replace($($CurrentGPU.'DriverDesc' | where {$_.Length -ne 0}), $($GPUDevice.AdapterCompatibility), "", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase);

# | Filter out proper memory size from regisrty cause registry query returns memory szie as Array with empty zero element, while integer required
'Memory' = "$(($CurrentGPU.'HardwareInformation.qwMemorySize' | where {$_.Length -ne 0}) / 1Mb) Mb";
}
}


# | Physical Adapters
foreach ($PhysicalAdapter in $PhysicalAdaptersQuery){

# Query Win32_NetworkAdapter to get more informative AdapterType string and MACAddress
$PhysicalAdapterInfo = Get-CimInstance -Query "Select PNPDeviceID, AdapterType, MACAddress, Manufacturer from Win32_NetworkAdapter" | where {$_.PNPDeviceID -like "$($PhysicalAdapter.PNPDeviceID)"}

# | Add to PCInfo using each PhysicalAdapter Array index to enumerate
$PCInfo["PhysicalNetAdapter$($PhysicalAdaptersQuery.IndexOf($PhysicalAdapter))"] = [ordered]@{

# | Manufacturer from Win32_NetworkAdapter
'Manufacturer' = $PhysicalAdapterInfo.Manufacturer;

# | Name from MSFT_NetAdapter cleaned of redundant manufacturer info
'Name' = [regex]::Replace($PhysicalAdapter.interfaceDescription, $PhysicalAdapterInfo.Manufacturer, "", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase);

# MAC from Win32_NetworkAdapter
'MACAddress' = $PhysicalAdapterInfo.MACAddress; 

# More Informative MediaType from Win32_NetworkAdapter
'AdapterType' = $PhysicalAdapterInfo.AdapterType;

# Speed as INT from MSFT_NetAdapter convert to Mbits
'AdapterSpeed' = "$($PhysicalAdapter.Speed / 1000000) Mbit";

}
}

# | IPEnabled Connections (Both Active and InActive)
foreach ($Connection in $NetConnectionsQuery) {

# | Add Current Connection to PCInfo using Array index for enumeration
$PCInfo["NetConnection$($NetConnectionsQuery.IndexOf($Connection))$($ConnectionStatus)"] = [ordered]@{

# | Connection Name from Win32_NetworkAdapter
'Name' = $Connection.Description;

# | Connection MAC Address
'MACAddress' = $Connection.MACAddress;

# | Use regexp to get only IPv4 Address
'IPv4Address' = [string]$Connection.IPAddress.split().where({$_ -match "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"})

# | Use regexp to get rid of metrics
'SubnetMask' = [string]$Connection.IPSubnet.split().where({$_ -match "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"})

# | Convert to string cause initially powershell returns Gateway as Array
'Gateway' = [string]$Connection.DefaultIpGateway;
}
}

# | Export to JSON
IF ($JSON -ne 0) {

# | Set File Extension
$FileExtension = "json"

# | Set Path and Filename
$LocalPath = "$FilePath\$PCName.$FileExtension"

# | If JSON file Already Exist
IF ([System.IO.File]::Exists($LocalPath)) {

# | Add New Line Separator just for convinience
Add-Content -Path $LocalPath -Value "----------New Entry----------"
}

# | Convert to JSON
$JSONFormat = $PCInfo | ConvertTo-JSON

# | Add to file
Add-Content -Path $LocalPath -Value $JSONFormat
}


# | # | Export to CSV
# | I am aware about ConvertTo-CSV and Export-CSV CMDlets yet in this case there is a need for more manageble CSV structure
# | Also by default Windows 10 PSVersion is 5, and Export-CSV -UseQuotes only available in PSVersion 7
# | since I need legacy OS support and CSV with quotes obligatory usage of Export-CSV, etc. somewhat pointless.

IF ($CSV -ne 0) {

# | Set File Extension
$FileExtension = "csv"

# | Set Path and Filename
$LocalPath = "$FilePath\$PCName.$FileExtension"

# | If CSV file Already Exist
IF ([System.IO.File]::Exists($LocalPath)) {

# | Add New Line Separator just for convinience
Add-Content -Path $LocalPath -Value "----------New Entry----------"

}

foreach ($Key in $PCInfo.keys) {

# | If Current element Type is String it means that current Dictionary is first level inlayed dictionary
IF ($PCInfo.$Key.GetType() -like "*String*") {

# | Add Dictionary Key and Key Value to *.csv file using ` symbol to ecape double quotes
Add-Content -Path $LocalPath -Value "`"$Key`"$StringCSVDelimiter`"$($PCInfo.$Key)`"" 

# | if Dictionary Type is Dictionary it means that I have second level inlayed dictionary
} ELSEIF ($PCInfo.$Key.GetType() -like "*Dictionary*") {

# | Add to file Dictionary key and values using join operator and "," as delimiter to form proper CSV structure
Add-Content -Path $LocalPath -Value `"$($Key, ($($PCInfo.$Key.Values) -join $DictionaryCSVDelimiter) -join $DictionaryCSVDelimiter)`"

}
}
}

# | Export to XML
# | I am aware about Export-Clixml and bunch of others XML and Clixml-Export cmdlets, yet I want readable structure with human understandable tags.
IF ($XML -ne 0) {

# | Set File Extension
$FileExtension = "xml"

# | Set Path and Filename
$LocalPath = "$FilePath\$PCName.$FileExtension"

# | Check if *.xml file has already exists and if NOT
IF (![System.IO.File]::Exists($LocalPath)) {

# | Write XML Notation
Add-Content -Path $LocalPath -Value '<?xml version="1.0" encoding="UTF-8"?>'

# Write XML Root Open Tag
Add-Content -Path $LocalPath -Value "<pcinfo>"
}

# | Get existing XML file content and rewrite without XML Close Root Tag
(Get-Content($LocalPath)).Replace("</pcinfo>", "") | Out-File -FilePath $LocalPath

# | Identify new info Begining
Add-Content -Path $LocalPath -Value "<!-- New Entry -->"

foreach ($Key in $PCInfo.Keys) {

# | Remove / and () from $Key, since they are not allowed in XML Tags
# | Also XML Tags must not contain spaces
$XMLKey = $Key -replace '[/()\[\]]', ""

# | If Element is String then I deal with Inlayed Dictionary of first level
IF ($PCInfo.$Key.GetType() -like "*String*") {

# | Write XMLKey as XML OpenTag, write Value, write /XMLKey as XML CloseTag
Add-Content -Path $LocalPath -Value $("`<$($XMLKey)`>$($PCInfo.$Key)`<`/$($XMLKey)`>")
}

# | If Element is second level Inlayed Dictionary
ELSEIF ($PCInfo.$Key -like "*Dictionary*") {

# | Write Inlayed Dictionary key as XML Open tag
Add-Content -Path $LocalPath -Value "`<$($XMLKey)`>"

# | Cycle through every value of inlayed Dictionary
foreach ($L2Dictionary in $PCInfo.$Key)
{

# | Use loop to iterate through "key:value" pairs in Dictionary
foreach ($SubKey in $L2Dictionary.Keys) {

# | Remove / and () from $SubKey, since they are not allowed in XML Tags
$XMLSubKey = $SubKey -replace '[/()\[\]]', ""

# | Write XMLSubKey name as XML Open tag, write key value, write /XMLSubKey name as XML Close tag
Add-Content -Path $LocalPath -Value $("`<$($XMLSubKey)`>$($L2Dictionary.$SubKey)`</$($XMLSubKey)`>") -NoNewline
}
}

# | Write inlayed Dictionary XMLKey as XML Close tag
Add-Content -Path $LocalPath -Value "`</$($XMLKey)`>"

}
}

# | Close XML root Tag
Add-Content -Path $LocalPath -Value "</pcinfo>"

}

# | Export to HTML
# | I am Aware about ConvertTo-HTML cmdlet, yet it is basic and use outdated table-style markup
# | Since I need div-tables it is easier to use custom HTML output cause parcing and editing ConvertTo-HTML output somewhat pointless
# | cause it takes more or same amount of code than simply write Main Dictionary content to html-file.
# | Also, for correct usage of ExportTo-HTML cmdlet I need to convert Main Hash-table to PSCustomObject and why would I want that
# | since it requires more resourses than simply write hash-table content to file.
# | Besides, I use HTML representation as GUI and it must have kind of NON-disgusting look, lol.

IF ($HTML -ne 0) {

# | Set File Extension
$FileExtension = "html"

# | Set Path and Filename
$LocalPath = "$FilePath\$PCName.$FileExtension"

# | Set our Beautiful Custom CSS Style
$CSSStyle = @"
<head><style>
.div-table {display: inline-block; width: 1200px; height: auto; min-height: 780px; border-left: 1px solid black; border-right: 1px solid black; padding: 10px; float: left; margin-top: 10px; margin-left: 10px; margin-bottom: 10px; vertical-align: top;}
.div-table-row {display: table-row;}
.div-table-cell {width: 180px; min-width: 180px; max-width: 180px; padding: 10px 20px 10px; border-top: 1px solid black; border-left: 1px solid black; float: left;}
</head></style>
"@

# | If HTML file Does NOT exist
IF (![System.IO.File]::Exists($LocalPath)) {

# | Write HTML Notation
Add-Content -Path $LocalPath -Value '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">'

# | Write CSS Style
Add-Content -Path $LocalPath -Value $CSSStyle
}

# | Open Main Div Table
Add-Content -Path $LocalPath -Value '<div class="div-table">'

# | Set Up Div Elements
$OpenRow = '<div class="div-table-row">'
$OpenDiv = '<div class="div-table-cell">'
$CloseDiv = $CloseRow = '</div>'

foreach ($Key in $PCInfo.Keys) {

# | If Level One Element Key is String (means it has no inlayed dictionaries)
IF ($PCInfo.$Key.GetType() -like "*String*") {

# | Add Dictionary Name as Row
Add-Content -Path $LocalPath -Value "$OpenRow $($Key) $CloseRow"

# | Add Dictionary Value as Div
Add-Content -Path $LocalPath -Value "$OpenRow $OpenDiv $($PCInfo.$Key) $CloseDiv $CloseRow"

# | If element is Dictionary which means it is Level One inlayed Dictionary
} ELSEIF ($PCInfo.$Key -like "*Dictionary*") {

# | Add Dictionary name as row header
Add-Content -Path $LocalPath -Value "$OpenRow $($Key) $CloseRow"

# | Cycle through every value of inlayed Dictionaries
foreach ($L2Dictionary in $PCInfo.$Key){

# | Open Row for several inllayed divs from hash values
Add-Content -Path $LocalPath -Value $OpenRow

foreach ($SubKey in $L2Dictionary.Keys) {

# | Write key name as Open tag, write key value, write key name as Close tag
Add-Content -Path $LocalPath -Value "$OpenDiv $($L2Dictionary.$SubKey) $CloseDiv" -NoNewline
}

# | Close Row With Inlayed Div Values from Inlayed Dictionary
Add-Content -Path $LocalPath -Value $CloseRow

}
}
}
# | Close Main Div Table
Add-Content -Path $LocalPath -Value $CloseRow
}
