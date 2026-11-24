@echo off
rem NetMon - create report for the last 7 days (edit -Days to change)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Make-Report.ps1" -Days 7
pause
