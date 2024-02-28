rem | Allow execution of hwinfo.ps1 (Administrator priviliges reuired)
powershell.exe -ExecutionPolicy Bypass -File .\hwinfo.ps1

rem | Start hwinfo.ps1 from 
powershell.exe -noexit -file "%~dp0hwinfo.ps1" 

rem | Restore Default Execution Policy
powershell Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Default
