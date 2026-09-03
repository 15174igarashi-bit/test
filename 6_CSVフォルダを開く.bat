@echo off
rem ------------------------------------------------------------------
rem  Open the destination folder (houjinkan kobetsu nagashikomi CSV).
rem ------------------------------------------------------------------
setlocal
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Open-KeiriMonth.ps1" -CsvFolder
if errorlevel 1 pause
endlocal
