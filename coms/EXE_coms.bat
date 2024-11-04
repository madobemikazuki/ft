@echo off
cd .\commands\ft_utils
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
PowerShell -command "(Measure-Command {. .\coms.ps1}).TotalSeconds"
PowerShell -NoExit -Nologo
