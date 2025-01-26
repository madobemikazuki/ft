@echo off
cd .\commands\tools\watchers
PowerShell Set-ExecutionPolicy RemoteSigned -Scope Process -Force
PowerShell -command . .\APD_watcher.ps1
PowerShell -NoExit -Nologo
