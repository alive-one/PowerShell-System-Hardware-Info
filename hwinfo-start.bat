rem | Allow execution of hwinfo.ps1 
powershell -ExecutionPolicy Bypass -File .\hwinfo.ps1

rem | After hwinfo.ps1 executed, return Default execution policy
powershell Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Default
