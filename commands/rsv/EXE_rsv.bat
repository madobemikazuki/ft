@echo off
cd .\commands\ft_utils
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force 
rem PowerShell -Windowstyle Hidden -NoProfile -File .\rsv.ps1
PowerShell -command "(Measure-Command {. .\rsv.ps1}).TotalSeconds"
PowerShell -NoExit -Nologo
