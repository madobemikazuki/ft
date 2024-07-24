@echo off
cd .\commands
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
@echo off
PowerShell clear
PowerShell . .\rsv.ps1
PowerShell . .\gZEN.ps1
PowerShell . .\bind_r.ps1
PowerShell . .\bind_c.ps1
PowerShell -NoExit -Nologo
