@echo off
cd .\commands\ft_utils
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
@echo off
PowerShell clear
PowerShell . .\rsv.ps1
PowerShell . .\gZEN.ps1
PowerShell . .\registered.ps1
PowerShell . .\coms.ps1
PowerShell . .\applicants.ps1
PowerShell -NoExit -Nologo
