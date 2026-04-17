@echo off
:: ============================================================
::  Export-ServerInfo.bat
::  Launcher for Export-ServerInfo.ps1
::
::  Double-click this file on the customer's server.
::  Output files are saved in the results\ subfolder.
:: ============================================================

cd /d "%~dp0"

echo.
echo  Starting server information export...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Export-ServerInfo.ps1"

if %ERRORLEVEL% neq 0 (
    echo.
    echo  ERROR: Script failed. Check logs\ folder for details.
    echo.
)

pause
