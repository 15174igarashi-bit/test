@echo off
rem ------------------------------------------------------------------
rem  Run once: put two shortcuts on the Desktop.
rem ------------------------------------------------------------------
setlocal
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\New-DesktopShortcut.ps1"
pause
endlocal
