@echo off
cd .\commands
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force 
PowerShell -command "(Measure-Command {. .\FnD.ps1}).TotalSeconds"
PowerShell -NoExit -Nologo
