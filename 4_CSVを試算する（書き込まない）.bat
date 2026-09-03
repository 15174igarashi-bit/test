@echo off
rem ------------------------------------------------------------------
rem  Dry run: build the CSV and show the checks WITHOUT touching file 2.
rem  Optional argument: 202610 / 2026-10 / -1 (previous month) / +1
rem ------------------------------------------------------------------
setlocal
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\New-YunyuFurikaeCsv.ps1" -Month "%~1" -DryRun
pause
endlocal
