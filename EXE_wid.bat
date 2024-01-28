@echo off
cd .\commands
Powershell -ExecutionPolicy RemoteSigned -Windowstyle Hidden -NoProfile -File .\wid.ps1
