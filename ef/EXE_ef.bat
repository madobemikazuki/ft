@echo off
cd .\commands\tools
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
PowerShell . .\ef.ps1
PowerShell -NoExit -Nologo
