@echo off
cd .\commands
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force 
@echo off
PowerShell clear
PowerShell -Windowstyle Hidden -NoProfile -File .\tools\coms.ps1
@echo off
PowerShell -Nologo
