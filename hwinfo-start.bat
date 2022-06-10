rem | For default powershell scripts execution policy is -Restricted
rem | First we need to allow execution of our *.ps1 script
powershell -ExecutionPolicy Bypass -File .\hwinfo.ps1
  
rem | When script is executed return default execution policy for security
powershell Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Default
  
pause
