@echo off
cd .\commands
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
@echo off
PowerShell clear
:input
PowerShell -command "(Measure-Command {. .\cnc.ps1}).TotalSeconds"
goto :input
