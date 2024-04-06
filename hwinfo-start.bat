rem | Allow execution of hwinfo.ps1 (Administrator priviliges reuired)
powershell.exe Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
  
rem | Start hwinfo.ps1 from
powershell.exe -noexit -file "%~dp0hwinfo.ps1"
  
rem | Restore Default Execution Policy
powershell.exe Set-ExecutionPolicy -Scope CurrentUser Default
