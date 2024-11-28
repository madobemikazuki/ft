@echo off
cd commands
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
PowerShell clear
:input
PowerShell -command "(Measure-Command {. .\ed.ps1}).TotalSeconds"
goto :input
