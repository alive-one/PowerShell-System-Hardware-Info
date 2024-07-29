# | ---CONFIGURATION---

# | Path to write data (Default is script root directory)
$FilePath = "$PSScriptRoot" 

# | $JSON = 1 to export Data in *.json format. (Multiple Choice is Possible)
$JSON = 0

# | $CSV = 1 to export Data in *.csv format.
$CSV = 0

# | Set standard CSV delimiter which is comma (,) by default
$StringCSVDelimiter = ','
$DictionaryCSVDelimiter = '","'

# | Yet some programs, such as MS Excel, use regional settings when import CSV Data
# | If you prefer to use your regional settings for CSV export, UNCOMMENT two lines below
#$StringCSVDelimiter = (Get-Culture).TextInfo.ListSeparator
#$DictionaryCSVDelimiter = "`"$($StringCSVDelimiter)`""

# | $XML = 1 to export Data in *.xml format.
$XML = 0

# | $HTML = 1 to export Data in *.html(GUI) format.
$HTML = 1

# | $SQL = 1 to export Data to MySQL Server.
# | Supposed that you already have MySQL Server up, running and properly setup
# | *.sql script to create database and setup user and user privileges you can download from github <github-link>
# | https://github.com/alive-one/PowerShell-System-Hardware-Info/blob/main/create_mysql_database.sql
$MySQL = 0

# | MySQL Server Address
$ServerIP = "192.168.0.70"

# | MySQL Server Port (Default is 3306)
$ServerPort = "3306"

# | Your Database Name here
$DatabaseName = "pcinfo"

# | MySQLServer username
$Username = "your-mysql-username"

# | SQL Server Password as Secure String
$Password = "your-secure-password" 




# --- DATA PROCESSING ---

# | As hwinfo.ps1 is a part of more complex system, current script must already have $PCName defined from previuos module "Autoren"
# | If not then this script is launched as standalone module and need to define $PCName as local System Name
IF ([string]::IsNUllOrEmpty($PCName)) {
[string]$PCName = $(hostname)
}

# | Ordered Dictionary to Store Data in strict order
$PCInfo = [ordered]@{

# | Local System Date (Be wary for it may differ from Accurate Date sometimes)
'LocalDate' = [string]$(Get-Date -Format "dd/MM/yyyy HH:mm"); 

# | Literally. 'Nuff said 
'SystemName' = $PCName;

}

# | Get RAM Slots Count in advance for I need it as part of Baseboard Info
# | All RAM Slots Available (No Matter Populaterd or Not)
# | Some baseboards and OS (especially server OS) return Memory Devices info not as [int] but as an [Array]
# | So next three lines of code check that output data is in correct format
$RAMSlotsCount = (Get-CimInstance -Query "Select MemoryDevices from Win32_PhysicalMemoryArray").MemoryDevices

if ([string]::IsNullOrEmpty($RAMSlotsCount)) {$RAMSlotsCount = "Unknown Slots Count"}

elseif ($RAMSlotsCount -is [System.Array] -and $RAMSlotsCount.Length -ge 0) {[string]$RAMSlotsCount = "$(($RAMSlotsCount | Measure -Sum).Sum) RAM Slots"}

else {$RAMSlotsCount = "$RAMSlotsCount RAM Slots"}




# | --- BASEBOARD INFO ---

# | WQL Query for Baseboard Info
$BaseBoardQuery = Get-CimInstance -Query "Select Manufacturer, Product from win32_Baseboard"

# | Check of we got something useful in our query
if ([string]::IsNullOrEmpty($BaseBoardQuery)) {

$BaseBoardManufacturer = "Noname"

$BaseBoardModel = "Unknown"

# | Check if necessary peoperties are defined
} else {

if ([string]::IsNullOrEmpty($BaseBoardQuery.Manufacturer)) {$BaseBoardManufacturer = "Noname"} else {$BaseBoardManufacturer = $BaseBoardQuery.Manufacturer}

if ([string]::IsNullOrEmpty($BaseBoardQuery.Product)) {$BaseBoardModel = "Unknown"} else {$BaseBoardModel = $BaseBoardQuery.Product}
}

# | Add data to $PCInfo
$PCInfo["Baseboard"] = [ordered]@{

'Manufacturer' = $BaseBoardManufacturer;

'Model' = $BaseBoardModel;

'RAMSlots' = $RAMSlotsCount;

}




# | --- CPU INFO ---

# | WQL Query for CPUs Socket, Name, etc.
$CPUQuery = @(Get-CimInstance -Query "Select DeviceID, Manufacturer, SocketDesignation, Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed from Win32_Processor")

# | Check If WQL Query return something useful or not
if ([string]::IsNullOrEmpty($CPUQuery)) {

$CPUEnumerator = "Unknown"

$CPUName = "Unknown"

$CPUManufacturer = "Unknown"

$CPUSocket = "Unknown"

$CPUSpeed  = "Unknown"

$CPUCores  = "Unknown"

$CPUThreads = "Unknown"

} else {

foreach ($CPU in $CPUQuery) {

if ([string]::IsNullOrEmpty($CPU.DeviceID)) {$CPUEnumerator = "Unknown$($CPUQuery.IndexOf($CPU))"} else {
$CPUEnumerator = $CPU.DeviceID
}

# | Check CPU Name string
if ([string]::IsNullOrEmpty($CPU.Name)) {$CPUName = "Unknown"} else {$CPUName = $CPU.Name}

# | Check CPU Manufacturer string
if ([string]::IsNullOrEmpty($CPU.Manufacturer)) {$CPUManufacturer = "Unknown"} else {$CPUManufacturer = $CPU.Manufacturer}

# | IF CPU Manufacturer is Intel or AMD
if ([string]$CPUManufacturer -match "Intel|AMD") {

# | Leave only Intel or AMD as Manufacturer String using predefined array Matches[] which contains results of last comparison
$CPUManufacturer = [string]$Matches[0]

# | Remove Manufacturer from CPU Name with regexp to get rid of data doubling;
$CPUName = [string]$CPUName -replace ".*?($CPUManufacturer).*? ", ""

} 

# | Check Socket Data
if ([string]::IsNullOrEmpty($($CPU.SocketDesignation))) {$CPUSocket = "Unknown"} else {$CPUSocket = $($CPU.SocketDesignation)}

# | Check Speed Data
if ([string]::IsNullOrEmpty($($CPU.MaxClockSpeed))) {$CPUSpeed = "Unknown"} else {$CPUSpeed = $($CPU.MaxClockSpeed)}

# | Check Cores Count 
if ([string]::IsNullOrEmpty($($CPU.NumberOfCores))) {$CPUCores = "Unknown"} else {$CPUCores = $($CPU.NumberOfCores)}

# | Check Threads Count
if ([string]::IsNullOrEmpty($($CPU.NumberOfLogicalProcessors))) {$CPUThreads = "Unknown"} else {$CPUThreads = $($CPU.NumberOfLogicalProcessors)}

# | Write CPU Data to $PCInfo
$PCInfo[$CPUEnumerator] = [ordered]@{

'Socket' = $CPUSocket;

'Manufacturer' = $CPUManufacturer;

'Name' = $CPUName;

'Speed' = "$CPUSpeed MHz";

'Cores/Threads' = "$($CPUCores) Cores / $($CPUThreads) Threads"

}

}

}




# | --- COMMON RAM INFO ---

# | Predefined RAM Codes from Win32_PhysicalMemory.MemoryType and SMBIOSMemoryType
$RAMTypes = @{0 = 'Unknown'; 1 = 'Other'; 20 = 'DDR'; 21 = 'DDR2'; 22 = 'DDR2 FB-DIMM'; 24 = 'DDR3'; 26 = 'DDR4'; 34 = 'DDR5'}

# | WQL Query for RAM Module Name, Size, Speed, MemoryType
$RAMQuery = @(Get-CimInstance -Query "Select DeviceLocator, BankLabel, Capacity, Speed, MemoryType, SMBIOSMemoryType from Win32_PhysicalMemory")

# | Check if RAMQuery contains something of use
if ([string]::IsNullOrEmpty($RAMQuery)) {

$RAMModuleEnumarator = "Unknown"

$RAMType = "Unknown"

$TotalRAMCapacity = "Unknown"

$RAMSpeed = "Unknown"

$DeviceLocator = "Unknown"

$RAMModuleCapacity = "Unknown"


} else {

# | Detecting RAM Type
# | RAM Type stored in MemoryType, SMBIOSMemoryTYpe or Both. So Use .where() method to Filter Out only first suitable value
[string]$RAMType = $($RAMTypes.[int]$RAMQuery.MemoryType[0], $RAMTypes.[int]$RAMQuery.SMBIOSMemoryType[0]).where({$_ -notlike 'Unknown'},'First')

# | Check if RAMType defined correctly
if ([string]::IsNullOrEmpty($RAMType)) {$RAMType = "Unknown"}

# | Calsulate All RAM Available
[string]$TotalRAMCapacity = "$(($RAMQuery.Capacity | Measure -Sum).Sum / 1Gb) Gb"

# | Check Total RAM Capacity
if ([string]::IsNullOrEmpty($TotalRAMCapacity)) {$TotalRAMCapacity = "Unknown"}

# | Check RAM Speed
if ([string]::IsNullOrEmpty($RAMQuery[0].Speed)) {$RAMSpeed = "Unknown"} else {$RAMSpeed = "$($RAMQuery[0].Speed) Mhz"}

# | Write Common RAM Info to $PCInfo Dictionary
$PCInfo['RAM'] = [ordered]@{

'RAMType' = $RAMType;

'TotalRAM' = $TotalRAMCapacity;

'RAMSpeed' = $RAMSpeed;

}

}




# | --- INDIVIDUAL RAMMODULES INFO ---

# | Checking and Cleaning RAM Modules Data
foreach ($RAMModule in $RAMQuery) {

# | If BankLabel unknown enumerate it with help of array index from RAMQuery
# | since if there are several RAMModeules without enumeration they owerwrte each other
if ([string]::IsNullOrEmpty($RAMModule.BankLabel)) { $RAMModuleEnumerator = "Unknown$($RAMQuery.IndexOf($RAMModule))" } else { 

# | Set RAMModuleEnumerator and remove whitespaces for XML does not allow whitespaces in tags
#$RAMModuleEnumerator = "RAM$($RAMModule.BankLabel).replace(" ", "")"
$RAMModuleEnumerator = ("RAM$($RAMModule.BankLabel)").replace(" ", "")

}

# | Check Slot Name
if ([string]::IsNullOrEmpty($RAMModule.DeviceLocator)) { $DeviceLocator = "Unknown" } else { $DeviceLocator = $RAMModule.DeviceLocator }

# | Check Module Capacity
if ([string]::IsNullOrEmpty($RAMModule.Capacity)) {$RAMModuleCapacity = "Unknown" } else { $RAMModuleCapacity = "$($RAMModule.Capacity /1Mb) Mb" }

# | Write RAM Module Info to $PCInfo
$PCInfo[$RAMModuleEnumerator] = [ordered]@{

'Slot' = $DeviceLocator;

'Size' = $RAMModuleCapacity;

}

}




# | --- STORAGE INFO ---

# | Physical Storage Devices excluding USB, SD-Cards, other Removables
# | Use Get-PhysicalDisk because only this cmdlet return proper values for BusType and MediaType, i.e. (SATA) and SSD accordingly
# | (All others, like Win32_DiskDrive or MSFT_PhysicalDisk lack some info, either storage type or bus type detected wrong or not present)
$StorageQuery = @(Get-PhysicalDisk | where {$_.BusType -notlike "USB" -xor "SD" -xor "MMC" -xor "Unknown" -xor "Virtual"} | Select DeviceID, BusType, MediaType, FriendlyName, Size, SerialNumber | Sort-Object -Property DeviceID)

# | Check if WQL Storage query return something useful
if ([string]::IsNullOrEmpty($StorageQuery)) {

$DeviceEnumerator = "Unknown"

$DevicePNP = "Unknown"

$DeviceManufacturer = "Unknown"

$DeviceBusType = "Unknown"

$DeviceMediaType = "Unknown"

$DeviceModel = "Unknown"

$DeviceSize = "Unknown"

} else {

# | For each Device in StorageQuery
foreach ($Device in $StorageQuery) {

# | From Win32_DiskDrive query PNPDeviceID, filtering by serial number to match Device which currently in cycle
$DevicePNP = (Get-CimInstance -Query "SELECT PNPDeviceID, SerialNumber FROM Win32_DiskDrive WHERE SerialNumber LIKE '%$($Device.SerialNumber)%'").PNPDeviceID

# | Check DevicePNP
if ([string]::IsNullOrEmpty($DevicePNP)) {

$DevicePNP = "Unknown"

$DeviceManufacturer = "Unknown"

} else {

# | Define and Check Device Manufacturer by getting rid of other stuff from DeviceID string
$DeviceManufacturer = [regex]::Match($DevicePNP, "VEN_(.+?)&").Groups[1].Value

}

# | if DeviceID not defined mark it as "unknown" and enumerate accordint to its array index to escape names conflict when add to $PCInfo
if ([string]::isNullOrEmpty($Device.DeviceID)) {$DeviceEnumerator = "Unknown$($StorageQuery.IndexOf($Device))"} else {$DeviceEnumerator = "Disk$($Device.DeviceID)"}

# | Check BusType
if ([string]::isNullOrEmpty($Device.BusType)) {$DeviceBusType = "Unknown"} else {$DeviceBusType = $Device.BusType}

# | Check MediaType
if ([string]::isNullOrEmpty($Device.MediaType)) {$DeviceMediaType = "Unknown"} else {$DeviceMediaType = $Device.MediaType}

# | Check Size
# | Use division by 1000000000 instead of Powershell 'Gb' to set 'proper' Disk Size
if ([string]::IsNullOrEmpty($Device.Size)) {$DeviceSize = "Unknown"} else {$DeviceSize = [string]([Math]::Round($Device.Size/1000000000)) + " Gb";}

# | Define and Check Model
# | Also get rid of Manufacturer name in Model name
if ([string]::isNullOrEmpty($Device.FriendlyName)) {$DeviceModel = "Unknown"} else {

# | Use regex to remove redundant info such as Manufacturer name from Model name
$DeviceModel = [regex]::Replace($Device.FriendlyName, $DeviceManufacturer, "", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase);
}


# | Add Data on Current Storage Device to $PCInfo
# | Use Each Storage Device Unique Identifier as enumerator to name Ordered Dictionary which will contain all current Storage Device info
$PCInfo[$DeviceEnumerator] = [ordered]@{

'(Bus)Type' = "($($DeviceBusType))$($DeviceMediaType)";

'Manufacturer' = $DeviceManufacturer;

'Model' = $DeviceModel;

'Size' = $DeviceSize;

}

}

}




# | --- GPUS INFO ---

# | GPU Name, DeviceID and Proper GPUMemory amount from Registry
$GPUsRegQuery = Get-ItemProperty -path "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4D36E968-E325-11CE-BFC1-08002BE10318}\????" | Select -Property MatchingDeviceID, DriverDesc, HardwareInformation.qwMemorySize 

# | DeviceIDs of all currently installed GPUs
# | Store GPU IDs Strings to Array to use Array Indexes for GPU devices enumeration later
# | Filter by PCI to get rid of Virtual Devices if any
$InstalledGPUs = @(Get-CIMInstance -Query "SELECT PNPDeviceID, AdapterCompatibility from Win32_VideoController WHERE PNPDeviceID LIKE 'PCI%'")

# | If there is no or insufficient data of GPUs from registry or WQL System Query
if ([string]::IsNullOrEmpty($GPUsRegQuery) -or [string]::IsNullOrEmpty($InstalledGPUs)) {

$GPUManufacturer = "Unknown"

$GPUModel = "Unknown"

$GPUMemory = "Unknown"

} else {

# | Cycle through every GPU    
foreach ($GPUDevice in $InstalledGPUs) {

# | Compare currently installed GPU Device IDs with corresponding values from Registry to get necessary registry branch with actual info on GPU Memory size
$CurrentGPU = $GPUsRegQuery | where {$GPUDevice.PNPDeviceID -like "$($_.MatchingDeviceID)*"}

#Check CurrentGPU
if ([string]::IsNullOrEmpty($CurrentGPU)) {

$GPUManufacturer = "Noname"

$GPUModel = "Noname"

$GPUMemory = "Unknown"
} else {

# | You can never have too many checks lol. Check GPU Manufacturer name
if ([string]::IsNullOrEmpty($GPUDevice.AdapterCompatibility)) {$GPUManufacturer = "Noname"} else {$GPUManufacturer = $GPUDevice.AdapterCompatibility}

# | Check GPU Name
if ([string]::IsNullOrEmpty($($CurrentGPU.'DriverDesc'))) {$GPUModel = "Noname"} else {

# | Remove Manufacturer Name from GPU Model Name since we have already store Manufacturer in appropriate separated field above
# | Also Trim() GPU Model string varuable for there is whitespace in the very begining
$GPUModel = ([regex]::Replace($($CurrentGPU.'DriverDesc' | where {$_.Length -ne 0}), $($GPUDevice.AdapterCompatibility), "", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)).Trim();
}

# | Check GPU Memory Size
if ([string]::IsNullOrEmpty($CurrentGPU.'HardwareInformation.qwMemorySize')) { $GPUMemory = "Unknown or DVMT" } else {

# | Filter out empty elements from Array with help to get proper memory size from regisrty cause registry query returns memory size as Array with empty zero element, while integer required
$GPUMemory = "$(($CurrentGPU.'HardwareInformation.qwMemorySize' | where {$_.Length -ne 0}) / 1Mb) Mb";
}

}

# | Use ArrayIndex from InstalledGPUDeviceIDs Query to create unique Dictionary name to store GPUName and Memory Size.
$PCInfo["GPU$($InstalledGPUs.IndexOf($GPUDevice))"] = [ordered]@{

# | GPU Manufacturer

'Manufacturer' = $GPUManufacturer;

'Model' = $GPUModel;

'Memory' = $GPUMemory
}

}

}




# | --- PHYSICAL ADAPTERS INFO ---

# | WQL Query Physical Adapters Only and store in Array to use Array Indexes for NetAdapters enumeration
$PhysicalAdaptersQuery = @(Get-CimInstance -ClassName MSFT_NetAdapter -Namespace ROOT\StandardCimv2 | Where-Object {!$_.Virtual})

# | Check if WQL Query contain some data or no
if ([string]::IsNullOrEmpty($PhysicalAdaptersQuery)) {

$AdapterEnumerator = "Unknown"

$PhysicalAdapterManufacturer = "Unknown"

$PhysicalAdapterModel = "Unknown"

$PhysicalAdapterMacAddress = "Unknown"

$PhysicalAdapterType = "Unknown"

$PhysicalAdapterSpeed = "Unknown"

} else {

# | Cycle through Physical Adapters in query
foreach ($PhysicalAdapter in $PhysicalAdaptersQuery){

# | Generate a name of ordered Dictionary to store detailes info on Adapter in main data set $PCInfo
$AdapterEnumerator = "PhysicalNetAdapter$($PhysicalAdaptersQuery.IndexOf($PhysicalAdapter))"

# Query Win32_NetworkAdapter to get more informative AdapterType string and MACAddress
$PhysicalAdapterInfo = Get-CimInstance -Query "Select PNPDeviceID, AdapterType, MACAddress, Manufacturer from Win32_NetworkAdapter" | where {$_.PNPDeviceID -like "$($PhysicalAdapter.PNPDeviceID)"}

# | Check Physical Adapter Manufacturer from Win32_NetworkAdapter
if ([string]::IsNullOrEmpty($PhysicalAdapterInfo.Manufacturer)) {$PhysicalAdapterManufacturer = "Unknown"} else {$PhysicalAdapterManufacturer = $PhysicalAdapterInfo.Manufacturer}

# | Check Physical Adapter Model
if ([string]::IsNullOrEmpty($PhysicalAdapter.InterfaceDescription)) {$PhysicalAdapterModel = "Unknown"} else {

# | Name from MSFT_NetAdapter cleaned of redundant manufacturer info and whitespaces
$PhysicalAdapterModel = ([regex]::Replace($PhysicalAdapter.InterfaceDescription, $PhysicalAdapterInfo.Manufacturer, "", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)).Trim();}

# | Check Physical Adapter MACAddress from Win32_NetworkAdapter
if ([string]::IsNullOrEmpty($PhysicalAdapterInfo.MACAddress)) {$PhysicalAdapterMacAddress = "Unknown"} else {$PhysicalAdapterMacAddress = $PhysicalAdapterInfo.MACAddress}

# | Check Physical Adapter Type from Win32_NetworkAdapter
if ([string]::IsNullOrEmpty($PhysicalAdapterInfo.AdapterType)) {$PhysicalAdapterType = "Unknown"} else {$PhysicalAdapterType = $PhysicalAdapterInfo.AdapterType}

# | Check Physical Adapter Speed from MSFT_NetAdapter and convert to Mbits if any
if ([string]::IsNullOrEmpty($PhysicalAdapter.Speed)) {$PhysicalAdapterSpeed = "Unknown"} else {$PhysicalAdapterSpeed = "$($PhysicalAdapter.Speed / 1000000) Mbit"}


# | Add Info to PCInfo
$PCInfo[$AdapterEnumerator] = [ordered]@{

'Manufacturer' = $PhysicalAdapterManufacturer;

'Model' = $PhysicalAdapterModel;

'MACAddress' = $PhysicalAdapterMACAddress; 

'AdapterType' = $PhysicalAdapterType;

'AdapterSpeed' = $PhysicalAdapterSpeed;

}

}

}




# | --- NETWORK CONNECTIONS INFO ---

# | Query Network Connections with IPEnbled 
$NetConnectionsQuery = @(Get-CimInstance -Query "Select Description, IPAddress, IPSubnet, DefaultIPGateway, MACAddress from Win32_NetworkAdapterConfiguration WHERE IPEnabled = 1")

# | Check if WQL Query returned something of value or not
if ([string]::IsNullOrEmpty($NetConnectionsQuery)) {

$ConnectionEnumerator = "Unknown"

$NetConnectionName = "Unknown"

$NetConnectionMacAddress = "Unknown"

$NetConnectionIPv4Address = "Unknown"

$NetConnectionSubnetMask = "Unknown"

$NetConnectionGateway = "Unknown"

} else {

# | IPEnabled Connections (Both Active and InActive)
foreach ($Connection in $NetConnectionsQuery) {

# | Enumerate connection according to its Array index
$ConnectionEnumerator = "NetConnection$($NetConnectionsQuery.IndexOf($Connection))"

# | Get Connection Name from Win32_NetworkAdapter using Win32_NetworkAdapter.DeviceID and Win32_NetworkAdapter.Index
# | as corresponding valies to filter out current connection Name
$NetConnectionName = (Get-CimInstance -Query "Select NetConnectionID from Win32_NetworkAdapter WHERE DeviceID = $($Connection.Index)").NetConnectionID

# | Check Connection Name
if ([string]::IsNullOrEmpty($NetConnectionName)) {$NetConnectionName = "Unknown"}

# | Check MAC Address
if ([string]::IsNullOrEmpty($Connection.MACAddress)) {$NetConnectionMacAddress = "Unknown"} else {$NetConnectionMacAddress = $Connection.MACAddress}

# | Check IPv4 Address
if ([string]::IsNullOrEmpty($Connection.IPAddress)) {$NetConnectionIPv4Address = "Unknown"} else {

# | Use regexp to get only IPv4 Address
$NetConnectionIPv4Address = [string]$Connection.IPAddress.split().where({$_ -match "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"})
}

# | Check Subnet Mask
if ([string]::IsNullOrEmpty($Connection.IPSubnet)) {$NetConnectionSubnetMask = "Unknown"} else {

# | Use regexp to get rid of metrics and such
$NetConnectionSubnetMask = [string]$Connection.IPSubnet.split().where({$_ -match "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b"})
}

# | Check Gateway
if ([string]::IsNullOrEmpty($Connection.DefaultIpGateway)) {$NetConnectionGateway = "Unknown"} else {

# | Convert to string cause initially powershell returns Gateway as Array
$NetConnectionGateway = [string]$Connection.DefaultIpGateway;
}

# | Add Current Connection to PCInfo using Array index for enumeration
$PCInfo[$ConnectionEnumerator] = [ordered]@{

'Name' = $NetConnectionName;

'MACAddress' = $Connection.MACAddress;

'IPv4Address' = $NetConnectionIPv4Address;

'SubnetMask' = $NetConnectionSubnetMask;

'Gateway' = $NetConnectionGateway;

}

}

}




# | --- EXPORT TO JSON ---

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

# | Export to CSV
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


# | Export to MySQL Server

IF ($MySQL -ne 0) {

# | Load MySQL\Connector (MySql.Data.dll) Assembly and dependencies
[Reflection.Assembly]::LoadFrom("$PSScriptRoot\MySql.Data.dll")
[Reflection.Assembly]::LoadFrom("$PSScriptRoot\System.Threading.Tasks.Extensions.dll")

# | Generate connection string and open Database connection
$ConnectionString = "server=$ServerIP;Port=$ServerPort;database=$DatabaseName;user id=$Username;password=$Password"
$connection = New-Object MySql.Data.MySqlClient.MySqlConnection($connectionString)
$connection.Open()

# | Select Database
$command = $connection.CreateCommand()
$command.CommandText = "use $DatabaseName;"
$command.ExecuteNonQuery() | Out-Null

# | Insert into BASEBOARDS table

#foreach ($BaseBoardItem in $PCInfo.Keys | where {$_ -match '^Baseboard\d+$'}) { - тут надо или менять структуру базы или придумать как несколько материнок связывать с прочей комплектухой, но скорее всего менять структуру базы

$command = $connection.CreateCommand()

# | Create Insert Command Text
$command.CommandText = "INSERT INTO baseboards (baseboard_manufacturer, baseboard_model, baseboard_ramslots) VALUES (@manufacturer, @model, @ramslots);"

# | Beware SQL-Injections! Use Prepared Statements. Переделать, чтобы матернки забиралдо из массвиа а не запроса
$command.Parameters.AddWithValue("@manufacturer", $($BaseBoardQuery.Manufacturer))
$command.Parameters.AddWithValue("@model", $($BaseBoardQuery.Product))
$command.Parameters.AddWithValue("@ramslots", $($PCInfo.Baseboard.RAMSlots -replace "[a-z]"))

# | Excute order 66! You know...
$command.ExecuteNonQuery() | Out-Null

# | Get the last inserted ID of baseboard. Note that there is no neeed to encapsulate connection in trasaction since LAST_INSERTED_ID() is Connection Specific
# | i.e. $BaseBoardID returns last id in range of current connection, not all database. When only operation is crating new record in DB it is fine.
$BaseBoardID = $command.LastInsertedId

# | For Each CPU Item in $PCInfo
foreach ($CPUItem in $PCInfo.Keys | where {$_ -match '^CPU\d+$'}) {

# | Strip current CPU 'Cores/Threads' from any alphabetical simbols and convert to Array containing number of Cores: CoresAndThreads[0] 
# | and number of Threads: CoresAndThreads[1] as zero and first element accordingly
$CoresAndThreads = $(($PCInfo.$CPUItem.'Cores/Threads' -replace "[a-z]") -split "/")

# | Insert into CPUS table
$command = $connection.CreateCommand()

# | Insert query with predefined values
$command.CommandText = "INSERT INTO cpus (cpu_socket, cpu_manufacturer, cpu_model, cpu_clock_mhz, cpu_cores, cpu_threads, baseboard_id) VALUES (@socket, @manufacturer, @model, @clockMHz, @cores, @threads, @baseboardId);"

# | Add values to query. Prepared Statements, yay! 
$command.Parameters.AddWithValue("@socket", $($PCInfo.$CPUItem.Socket -replace "[a-z]"))
$command.Parameters.AddWithValue("@manufacturer", $($PCInfo.$CPUItem.Manufacturer))
$command.Parameters.AddWithValue("@model", $($PCInfo.$CPUItem.Name))
$command.Parameters.AddWithValue("@clockMHz", $($PCInfo.$CPUItem.Speed -replace "[a-z]"))
$command.Parameters.AddWithValue("@cores", $($CoresAndThreads[0]))
$command.Parameters.AddWithValue("@threads", $($CoresAndThreads[1]))
$command.Parameters.AddWithValue("@baseboardId", $($BaseBoardID))

# | Execute command
$command.ExecuteNonQuery() | Out-Null
}

# | Insert into RAMTYPES table
$command = $connection.CreateCommand()
$command.CommandText = "INSERT INTO ramtypes (ram_type) VALUES (@ramtype);"
$command.Parameters.AddWithValue("@ramtype", $RAMType)
$command.ExecuteNonQuery() | Out-Null

# | Get RAMTypeID
$RAMTypeID = $command.LastInsertedId

# | Insert into RAMMODULES table
# | Note that RamModule Speed taken from RAMQuery directly
foreach ($RAMModuleItem in $PCInfo.Keys | where {$_ -match '^RAMBANK\d+$'}) {

$command = $connection.CreateCommand()
# | Create Prepared Statement
$command.CommandText = "INSERT INTO rammodules (bank_label, module_size_mb, module_speed_mhz, baseboard_id, ramtype_id) VALUES (@banklabel, @modulesize, @modulespeed, @baseboardId, @ramtypeId);"

# | Add Values
$command.Parameters.AddWithValue("@banklabel", $RamModuleItem)
$command.Parameters.AddWithValue("@modulesize", $($PCInfo.$RAMModuleItem.Size -replace "[a-z]"))
$command.Parameters.AddWithValue("@modulespeed", $($PCInfo.RAM.RAMSpeed -replace "[a-z]"))
$command.Parameters.AddWithValue("@baseboardId", $BaseBoardID)
$command.Parameters.AddWithValue("@ramtypeId", $RAMTypeID)

# | Execute insert command
$command.ExecuteNonQuery() | Out-Null
}

# | Insert into COMPUTERS table
$command = $connection.CreateCommand()

# | Note that serverdate is not part of Prepared Statement but assigned directlry from query
$command.CommandText = "INSERT INTO computers (computer_name, server_date, total_ram_gb, baseboard_id, ramtype_id) VALUES (@computername, NOW(), @totalram, @baseboardId, @ramtypeId);"

# Add Values
$command.Parameters.AddWithValue("@computername", $PCInfo.SystemName)
$command.Parameters.AddWithValue("@totalram", $($PCInfo.RAM.TotalRAM -replace "[a-z]"))
$command.Parameters.AddWithValue("@baseboardId", $BaseBoardID)
$command.Parameters.AddWithValue("@ramtypeId", $RAMTypeID)

# | Execute command
$command.ExecuteNonQuery() | Out-Null

# | Get ComputerID from computers table to use as foreign_key for each storage device
$ComputerID = $command.LastInsertedId

# | Insert into STORAGES table
foreach ($StorageItem in $PCInfo.Keys | where {$_ -match '^DISK\d+$'}) {

# | Check first if there are correct Bus and Type info
if ([string]::IsNullOrEmpty($PCInfo.$StorageItem.'(Bus)Type')) {

$BusAndType = @("Unknown", "Unknown")

} else {

# | Separate (Bus)Type to Bus and Type
$BusAndType = $($PCInfo.$StorageItem.'(Bus)Type' -split "\)")
}

$command = $connection.CreateCommand()

# | Prepared Statement
$command.CommandText = "INSERT INTO storages (storage_manufacturer, storage_model, storage_bus, storage_type, storage_size_gb, computer_id) VALUES (@storagemanufacturer, @storagemodel, @storagebus, @storagetype, @storagesizeb, @computerId);"

# | Adding Values
$command.Parameters.AddWithValue("@storagemanufacturer", $PCInfo.$StorageItem.Manufacturer)
$command.Parameters.AddWithValue("@storagemodel", $PCInfo.$StorageItem.Model)
$command.Parameters.AddWithValue("@storagebus", $($BusAndType[0] -replace "\(", ""))
$command.Parameters.AddWithValue("@storagetype", $BusAndType[1])
$command.Parameters.AddWithValue("@storagesizeb", $($PCInfo.$StorageItem.Size -replace "[a-z]"))
$command.Parameters.AddWithValue("@computerId", $ComputerID)

# | Execute command
$command.ExecuteNonQuery() | Out-Null
}

# | Insert into GPUS table
foreach ($GPUItem in $PCInfo.Keys | where {$_ -match '^GPU\d+$'}) {

$command = $connection.CreateCommand()
$command.CommandText = "INSERT INTO gpus (gpu_manufacturer, gpu_model, gpu_memory_mb, computer_id) VALUES (@gpumanufacturer, @gpumodel, @gpumemory, @computerId);"
$command.Parameters.AddWithValue("@gpumanufacturer", $PCInfo.$GPUItem.Manufacturer) 
$command.Parameters.AddWithValue("@gpumodel", $PCInfo.$GPUItem.Model)
$command.Parameters.AddWithValue("@gpumemory", $($PCInfo.$GPUItem.Memory -replace "[a-z]"))
$command.Parameters.AddWithValue("@computerId", $ComputerID)

# | Execute command
$command.ExecuteNonQuery() | Out-Null
}

# | Insert into NETADAPTERS table
foreach ($NetAdapterItem in $PCInfo.Keys | where {$_ -match '^PhysicalNetAdapter\d+$'}) {

$command = $connection.CreateCommand()
# | Prepared Statment
$command.CommandText = "INSERT INTO netadapters (netadapter_manufacturer, netadapter_model, netadapter_mac, netadapter_type, netadapter_speed_mbit, computer_id) VALUES (@netadaptermanufacturer, @netadaptermodel, @netadaptermac, @netadaptertype, @netadapterspeed, @computerId);" 

# | Add values to prepared statement
$command.Parameters.AddWithValue("@netadaptermanufacturer", $PCInfo.$NetAdapterItem.Manufacturer)
$command.Parameters.AddWithValue("@netadaptermodel", $PCInfo.$NetAdapterItem.Model)
$command.Parameters.AddWithValue("@netadaptermac", $PCInfo.$NetAdapterItem.MACAddress)
$command.Parameters.AddWithValue("@netadaptertype", $PCInfo.$NetAdapterItem.AdapterType)
$command.Parameters.AddWithValue("@netadapterspeed", $($PCInfo.$NetAdapterItem.AdapterSpeed -replace "[a-z]"))
$command.Parameters.AddWithValue("@computerId", $ComputerID)

# | Execute, yeah
$command.ExecuteNonQuery() | Out-Null
}

# | Insert into NETCONNECTIONS 
# | USE INET_ATON() function to convert and store network info as BIGINT UNSIGNED for it is MySQL recommended way to store IP Addresses and such
# | (To convert it back from BIGINT to IP-Addresses and other network stuff use MySQL INET_NTOA() function)
foreach ($NetConnectionItem in $PCInfo.Keys | where {$_ -match '^NetConnection\d+$'}) {
$command = $connection.CreateCommand()

# | Prepared Statement
$command.CommandText = "INSERT INTO netconnections (сonnection_name, connection_mac, connection_ip, connection_netmask, connection_gateway, computer_id) VALUES (@connectionname, @connectionmac, INET_ATON(@connectionip), INET_ATON(@connectionnetmask), INET_ATON(@connectiongateway), @computerId);"

# | Add Values
$command.Parameters.AddWithValue("@connectionname", $PCInfo.$NetConnectionItem.Name)
$command.Parameters.AddWithValue("@connectionmac", $PCInfo.$NetAdapterItem.MACAddress)
$command.Parameters.AddWithValue("@connectionip", $PCInfo.$NetConnectionItem.IPv4Address)
$command.Parameters.AddWithValue("@connectionnetmask", $PCInfo.$NetConnectionItem.SubnetMask)
$command.Parameters.AddWithValue("@connectiongateway", $PCInfo.$NetConnectionItem.Gateway)
$command.Parameters.AddWithValue("@computerId", $ComputerID)

# | Execute insert command
$command.ExecuteNonQuery() | Out-Null
}

# | Close Connection to MySQL Server
$connection.Close()

}
