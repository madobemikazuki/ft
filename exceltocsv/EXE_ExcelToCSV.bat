@echo off
cd .\commands
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
@echo off
PowerShell clear
PowerShell . .\ExcelToCSV.ps1
@echo off
PowerShell -NoExit -Nologo

