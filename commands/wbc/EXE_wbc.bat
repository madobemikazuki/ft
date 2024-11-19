@echo off
cd .\commands
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
@echo off
PowerShell clear
:input
PowerShell . .\wbc.ps1
@echo off
goto :input
rem PowerShell -NoExit -Nologo
