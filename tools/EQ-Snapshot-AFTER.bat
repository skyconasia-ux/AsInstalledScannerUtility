@echo off
REM Take AFTER snapshot -- run this AFTER making changes in the Web UI.
REM Output: C:\Temp\EQ_Snapshots\AFTER\
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0EQ-Snapshot.ps1" -Label AFTER
pause
