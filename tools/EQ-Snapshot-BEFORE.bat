@echo off
REM Take BEFORE snapshot -- run this BEFORE making changes in the Web UI.
REM Output: C:\Temp\EQ_Snapshots\BEFORE\
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0EQ-Snapshot.ps1" -Label BEFORE
pause
