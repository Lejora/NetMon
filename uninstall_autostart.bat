@echo off
rem NetMon - remove automatic start
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0NetMon.ps1" -UninstallStartup
pause
