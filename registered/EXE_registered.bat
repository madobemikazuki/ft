@echo off
cd .\commands\ft_utils
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
PowerShell . .\registered.ps1
PowerShell -NoExit -Nologo
