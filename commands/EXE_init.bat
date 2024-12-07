@echo off
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
cd .\commands\tools
PowerShell . .\dl.ps1
cd ..\ft_utils
PowerShell clear
PowerShell . .\rsv.ps1
PowerShell . .\gZEN.ps1
PowerShell . .\registered.ps1
PowerShell . .\coms.ps1
PowerShell . .\applicants.ps1
PowerShell -NoExit -Nologo
