@echo off
REM Export all Equitrac/ControlSuite configuration to a readable text file.
REM Output: results\EquitracConfig_HOSTNAME_DATE.txt  (next to this script)
REM
REM Requirements:
REM   - sqlite3.exe at C:\Windows\System32\sqlite3.exe
REM   - bcp.exe in PATH (SQL Server tools)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Export-EquitracConfig.ps1"
pause
