@echo off
rem PC health check: finds why the PC is slow or unstable and what to do.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test-PCHealth.ps1"
if errorlevel 1 pause
