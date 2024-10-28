@echo off
cd .\commands\ft_utils
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
@echo off
PowerShell clear
PowerShell . .\coh.ps1
PowerShell -NoExit -Nologo
