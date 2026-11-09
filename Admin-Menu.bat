@echo off
rem Main menu of all admin tools.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Admin-Menu.ps1"
if errorlevel 1 pause
