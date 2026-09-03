@echo off
rem ------------------------------------------------------------------
rem  Build the CSV from file 1 and overwrite file 2 (asks for Y/N).
rem  Optional argument: 202610 / 2026-10 / -1 (previous month) / +1
rem ------------------------------------------------------------------
setlocal
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\New-YunyuFurikaeCsv.ps1" -Month "%~1"
pause
endlocal
