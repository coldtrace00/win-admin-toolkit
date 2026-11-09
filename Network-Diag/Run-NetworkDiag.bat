@echo off
rem Network diagnostics: report is saved to the Desktop.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test-Network.ps1"
if errorlevel 1 pause
