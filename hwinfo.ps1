# | Configuration
# | Path to write data (Default is script root directory)
$FilePath = "$PSScriptRoot" 

# | Set Data Format Output. For example $JSON = 1 to export Data in *.json format. Multiple Choice is Possible
$JSON = 0

# | CSV Export Settings
$CSV = 0

# | Set standard CSV delimiter which is comma (,) by default
$StringCSVDelimiter = ','
$DictionaryCSVDelimiter = '","'

# | Yet some programs, such as MS Excell, use regional settings when import CSV Data
# | If you prefer to use your regional settings for CSV export, UNCOMMENT two lines below
#$StringCSVDelimiter = (Get-Culture).TextInfo.ListSeparator
#$DictionaryCSVDelimiter = "`"$($StringCSVDelimiter)`""

# | XML Export Settings
$XML = 0

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
'LocalDate' = [string]$(Get-Date -Format "dd/MM/yy HH:mm"); 

# | Literally. 'Nuff said 
'SystemName' = $PCName;
}

# | Use Queres where it is possbile because queries are fastest method to get CIM-Instance Data
# | Baseboard Manufacturer and Model
$BaseBoardQuery = Get-CimInstance -Query "Select Manufacturer, Product from win32_Baseboard"

# | Socket, Name, etc. CPU staff
$CPUQuery = Get-CimInstance -Query "Select DeviceID, SocketDesignation, Name, NumberOfCores, NumberOfLogicalProcessors from Win32_Processor"

# | Slot Name, Size, Speed, MemoryType
$RAMQuery = Get-CimInstance -Query "Select  DeviceLocator, Capacity, Speed, MemoryType, SMBIOSMemoryType from Win32_PhysicalMemory"

# | RAM Type stored in MemoryType, SMBIOSMemoryTYpe or Both. So Use .where() method to Filter Out only first suitable value
[string]$RAMType = $($RAMTypes.[int]$RAMQuery.MemoryType[0], $RAMTypes.[int]$RAMQuery.SMBIOSMemoryType[0]).where({$_ -notlike 'Undefined'},'First')

# | All RAM Slots Available (No Matter Populaterd or Not)
[string]$RAMSlotsCount = $(Get-CimInstance -Query "Select MemoryDevices from Win32_PhysicalMemoryArray").MemoryDevices

# | Physical Storage Devices excluding USB, SD-Cards, other Removables
$StorageQuery = Get-PhysicalDisk | where {$_.BusType -notlike "USB" -xor "SD" -xor "MMC" -xor "Unknown" -xor "Virtual"} | Select DeviceID, BusType, MediaType, FriendlyName, Size, HealthStatus | Sort-Object -Property DeviceID

# | GPU Name, DeviceID and Proper GPUMemory amount from Registry
$GPUsRegQuery = Get-ItemProperty -path "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4D36E968-E325-11CE-BFC1-08002BE10318}\????" | Select -Property MatchingDeviceID, DriverDesc, HardwareInformation.qwMemorySize 

# | DeviceIDs of all currently installed GPUs
# | Store GPU IDs Strings to Array to use Array Indexes for GPU devices enumeration later
$InstalledGPUDeviceIDs = @((Get-CIMInstance -Query "SELECT PNPDeviceID from Win32_VideoController WHERE PNPDeviceID LIKE 'PCI%'").PNPDeviceID)

# | Query All Network Adapters
$AllNetAdaptersQuery = Get-CIMInstance -Query "SELECT Name, MACAddress, AdapterType, NetConnectionID, NetConnectionStatus, PhysicalAdapter from Win32_NetworkAdapter"

# | Query Physical Adapters Only
$PhysicalAdaptersQuery = @($AllNetAdaptersQuery | where {$_.PhysicalAdapter -eq 1})

# | Query Network Connections with IPEnbled (Both Physical And Virtual)
$NetConnectionsQuery = @(Get-CimInstance Win32_NetworkAdapterConfiguration | where {$_.IPEnabled -eq 1} | Select IPAddress, IPSubnet, DefaultIPGateway, MACAddress)

# | Add Info to PCInfo Dictionary
# | Baseboard 
$PCInfo["BaseBoard"] = [ordered]@{
'Manufacturer' = $BaseBoardQuery.Manufacturer; 
'Model' = $BaseBoardQuery.Product; 
'RAMSLots' = "$($RAMSlotsCount) RAM Slots";
}

# | CPUs (There can be a few, you know)
foreach ($CPU in $CPUQuery) {

# | Use each CPU.DeviceID as name for Ordered Dictionary which contains CPU Name and Cores/Threads Count. And add this dictionary to PCInfo
$PCInfo[$CPU.DeviceID] = [ordered]@{
'Socket' = $CPU.SocketDesignation;
'Name' = $CPU.Name; 
'Cores/Threads' = "$($CPU.NumberOfCores) Cores / $($CPU.NumberOfLogicalProcessors) Threads";

}
}

# | RAM Type
$PCInfo['RAMType'] = $RAMType

# | RAM Modules
foreach ($RAMModule in $RAMQuery){
$PCInfo[$RAMModule.DeviceLocator] = [ordered]@{
'Size' = "$($RAMModule.Capacity / 1MB) Mb"; 
'Speed' = "$($RAMModule.Speed) Mhz";
}
} 

# | Storage
foreach ($Device in $StorageQuery) {

# | Use Each Storage Device Unique Identifier as enumerator to name Ordered Dictionary which contains all current Storage Device info
$PCInfo["Disk$($Device.DeviceID)"] = [ordered]@{

# | Define Bus and Storage Type
'(Bus)Type' = "($($Device.BusType))$($Device.MediaType)";

# | Literally Disk Friandly Name
'Name' = $Device.FriendlyName; 

# | Use division by 1000000000 instead of Powershell 'Gb' to set 'proper' Disk Size
'Size' = [string]([Math]::Round($Device.Size/1000000000)) + " Gb";

# | Windows Disk Status
'HealthStatus' = $Device.HealthStatus;
}
}

# GPUs
foreach ($DeviceID in $InstalledGPUDeviceIDs) {

# | Compare currently installed GPU Device IDs with corresponding values from Registry to get necessary registry branch with actual info on GPU Memory size
$CurrentGPU = $GPUsRegQuery | where {$DeviceID -like "$($_.MatchingDeviceID)*"}

# | Use ArrayIndex from InstalledGPUDeviceIDs Query to create unique Dictionary name to store GPUName and Memory Size. Then Add to PCInfo
$PCInfo["GPU$($InstalledGPUDeviceIDs.IndexOf($DeviceID))"] = [ordered]@{'Name' = [string]$($CurrentGPU.'DriverDesc') 

# | Convert from HEX registry value to suitable decimal view with help of .NET Method .ToDecimal
'Memory' = ([string](([System.Convert]::ToDecimal("$($CurrentGPU.'HardwareInformation.qwMemorySize')")) / 1MB) + " Mb") 
}
}

# | Physical Adapters
foreach ($PhysicalAdapter in $PhysicalAdaptersQuery){

# | Add to PCInfo using each PhysicalAdapter Array index to enumerate
$PCInfo["PhysicalAdapter$($PhysicalAdaptersQuery.IndexOf($PhysicalAdapter))"] = [ordered]@{

'Name' = [string]$PhysicalAdapter.Name; 

'MACAddress' = [string]$PhysicalAdapter.MACAddress; 

'AdapterType' = [string]$PhysicalAdapter.AdapterType;
}
}

# | IPEnabled Connections (Both Active and InActive)
foreach ($Connection in $NetConnectionsQuery) {

# | Connection Details
$ConnectionDetails = $AllNetAdaptersQuery | where {$_.MACAddress -like $Connection.MACAddress}

# Define Current NetConnection status (Active or InActive)
IF ($ConnectionDetails.NetConnectionStatus -eq 2) {$ConnectionStatus = "(Active)"} ELSE {$ConnectionStatus = "(InActive)"}

# | Add Current Connection to PCInfo
# | Use Array index of connection for enumeration
$PCInfo["NetConnection$($NetConnectionsQuery.IndexOf($Connection))$($ConnectionStatus)"] = [ordered]@{

# | Connection Name
'Name' = $ConnectionDetails.NetConnectionID;

# | Connection MAC Address
'MACAddress' = $ConnectionDetails.MACAddress;

# | Use regexp to get only IPv4 Address
'IPv4Address' = [string]$Connection.IPAddress.split().where({$_ -match "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"})

# | Use regexp to get rid of metrics
'SubnetMask' = [string]$Connection.IPSubnet.split().where({$_ -match "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"})

# | Convert to string cause initially powershell returns Gateway as Array
'Gateway' = [string]$Connection.DefaultIpGateway;
}
}

# | Export to JSON
IF ($JSON -eq 1) {

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

IF ($CSV -eq 1) {

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
IF ($XML -eq 1) {

# | Set File Extension
$FileExtension = "xml"

# | Set Path and Filename
$LocalPath = "$FilePath\$PCName.$FileExtension"

# | Check if *.xml file has already exists and if NOT
IF (![System.IO.File]::Exists($LocalPath)) {

# | Write XML Notation
Add-Content -Path $LocalPath -Value '<?xml version="1.0" encoding="UTF-8"?>'
}

# Write XML root Tag
Add-Content -Path $LocalPath -Value "<PCINFO>"

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
Add-Content -Path $LocalPath -Value "</PCINFO>"
}

# | Export to HTML
# | I am Aware about ConvertTo-HTML cmdlet, yet it is basic and use outdated table-style markup
# | Since I need div-tables it is easier to use custom HTML output cause parcing and editing ConvertTo-HTML output somewhat pointless
# | cause it takes more or same amount of code than simply write Main Dictionary content to html-file.
# | Also, for correct usage of ExportTo-HTML cmdlet I need to convert Main Hash-table to PSCustomObject and why would I want that
# | since it requires more resourses than simply write hash-table content to file.
# | Besides, I use HTML representation as GUI and it must have kind of NON-disgusting look, lol.

IF ($HTML -eq 1) {

# | Set File Extension
$FileExtension = "html"

# | Set Path and Filename
$LocalPath = "$FilePath\$PCName.$FileExtension"

# | Set our Beautiful Custom CSS Style
$CSSStyle = @"
<head><style>
.div-table {display: inline-block; width: 800px; height: auto; min-height: 780px; border-left: 1px solid black; border-right: 1px solid black; padding: 10px; float: left; margin-top: 10px; margin-left: 10px; margin-bottom: 10px; vertical-align: top;}
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
