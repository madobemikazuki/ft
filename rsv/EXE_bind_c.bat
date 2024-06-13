@echo off
cd .\commands
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
@echo off
PowerShell clear
PowerShell . .\bind_c.ps1
@echo off
PowerShell -NoExit -Nologo

