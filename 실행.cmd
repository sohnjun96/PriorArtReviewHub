@echo off
setlocal
powershell.exe -NoLogo -NoProfile -STA -File "%~dp0App\Start.ps1" -Demo
if errorlevel 1 pause
