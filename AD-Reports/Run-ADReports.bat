@echo off
rem Active Directory reports with a menu.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Get-ADReports.ps1"
if errorlevel 1 pause
