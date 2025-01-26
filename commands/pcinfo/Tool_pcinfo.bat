@echo off
cd .\commands\tools
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
PowerShell . .\pcinfo.ps1
PowerShell -NoExit -Nologo
