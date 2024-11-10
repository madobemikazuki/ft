@echo off
cd .\commands\tools
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
:input
PowerShell . .\ef.ps1
goto :input
