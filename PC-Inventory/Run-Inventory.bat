@echo off
rem Runs Get-PCInventory.ps1 from the same folder and saves an HTML report to the Desktop.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Get-PCInventory.ps1"
if errorlevel 1 pause
