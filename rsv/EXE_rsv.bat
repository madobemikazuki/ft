@echo off
cd .\commands\ft_utils
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force 
@echo off
PowerShell clear
rem PowerShell -Windowstyle Hidden -NoProfile -File .\rsv.ps1
PowerShell . .\rsv.ps1
PowerShell -NoExit -Nologo
