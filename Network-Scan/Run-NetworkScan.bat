@echo off
rem Network scanner: devices, open ports, risks and how to close them.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scan-Network.ps1"
if errorlevel 1 pause
