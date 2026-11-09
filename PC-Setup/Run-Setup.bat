@echo off
rem New PC setup: installs apps via winget and applies Windows settings (asks for admin rights).
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Setup-NewPC.ps1"
if errorlevel 1 pause
