@echo off
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force 
@echo off
PowerShell clear
PowerShell -Windowstyle Hidden -NoProfile -File .\commands\tools\coms.ps1
@echo off
PowerShell -Nologo