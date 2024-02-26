# PowerShell-System-Hardware-Info
PowerShell script to collect major hardware and some software information for local system. 
Designed for Windows OS family with Powershell version 5.1 or higher.
Output data formats are: JSON, CSV, XML or HTML(GUI). Multiple choise is possible.

Incase your Windows OS Powershell execution policy is "Restricted" (Which is highly possible cause its by Default) 
you need to change execution policy either manually or start hwinfo.ps1 via hwinfo-start.bat 
It will set *.ps1 files execution policy to "Bypass" and return it to "Default" later.
