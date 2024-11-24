@echo off
cd .\commands\tools
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force 
PowerShell -command "(Measure-Command {. .\FnD.ps1}).TotalSeconds"
PowerShell -NoExit -Nologo
