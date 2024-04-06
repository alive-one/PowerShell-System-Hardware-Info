# PowerShell-System-Hardware-Info
PowerShell script to collect major hardware and some software information for local system. 
Designed for Windows OS family with Powershell version 5.1 or higher.

When executed, script collects system hardware data, network settings, OS version and save as file in script's root directory using local system name as filename.
Output data formats are: JSON, CSV, XML or HTML(GUI). Multiple choiсe is possible. 
So, if you started script from D:\ drive, your system name is I123456-Home and choose *.csv as output file format you end up with file D:\I123456-Home.csv

Incase your Windows OS Powershell execution policy is "Restricted" (Which is highly possible because this policy enabled by Default) 
you need to change execution policy either manually or start hwinfo.ps1 via hwinfo-start.bat (Run as Administrator)
It will set *.ps1 files execution policy to "Bypass" and return it to "Default" later.
