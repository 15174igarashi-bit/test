@echo off
rem ------------------------------------------------------------------
rem  Run the self tests against a mock Box tree. Touches nothing real.
rem  Optional argument: path to a real workbook to test CSV generation.
rem ------------------------------------------------------------------
setlocal
chcp 65001 >nul
if "%~1"=="" (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0tests\Test-KeiriTools.ps1"
) else (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0tests\Test-KeiriTools.ps1" -Workbook "%~1"
)
pause
endlocal
