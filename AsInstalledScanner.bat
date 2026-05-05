@echo off
:: AsInstalledScanner.bat
:: Launcher for AsInstalledScanner.ps1
:: Usage: double-click for interactive menu, or pass a mode:
::   AsInstalledScanner.bat Before
::   AsInstalledScanner.bat After
::   AsInstalledScanner.bat Compare
::   AsInstalledScanner.bat Full

setlocal
set "SCRIPT=%~dp0tools\AsInstalledScanner.ps1"

if "%~1"=="" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Mode %~1
)

endlocal
