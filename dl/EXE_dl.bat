@echo off
cd .\commands\tools
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
@echo off
PowerShell clear
PowerShell  -command "(Measure-Command {. .\dl.ps1}).TotalSeconds"
PowerShell -NoExit -Nologo