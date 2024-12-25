@echo off
cd .\commands\tools
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
PowerShell . .\ezy_watch_commander.ps1
PowerShell -NoExit -Nologo
