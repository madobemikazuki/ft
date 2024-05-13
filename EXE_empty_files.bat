@echo off
cd commands
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
@echo off
PowerShell clear
PowerShell . .\tools\ef.ps1
@echo off
PowerShell -NoExit -Nologo
