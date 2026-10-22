@echo off
setlocal DisableDelayedExpansion
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File "%~dp0App\Start.ps1" -Demo -AdminMode
if errorlevel 1 pause
