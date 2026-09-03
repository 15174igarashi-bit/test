@echo off
rem ------------------------------------------------------------------
rem  Open this month's "houjinkan shori data (copy)" folder in Explorer.
rem  Optional argument: 202610 / 2026-10 / -1 (previous month) / +1
rem ------------------------------------------------------------------
setlocal
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\Open-KeiriMonth.ps1" -Month "%~1"
if errorlevel 1 pause
endlocal
