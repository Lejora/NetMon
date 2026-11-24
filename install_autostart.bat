@echo off
rem NetMon - start automatically (minimized) at Windows logon
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0NetMon.ps1" -InstallStartup
pause
