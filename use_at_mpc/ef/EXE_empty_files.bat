@echo off
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
@echo off
PowerShell clear
PowerShell . .\ef.ps1
@echo off
PowerShell -NoExit -Nologo