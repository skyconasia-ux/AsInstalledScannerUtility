@echo off
REM Compare BEFORE and AFTER snapshots.
REM Reads:  C:\Temp\EQ_Snapshots\BEFORE\  and  AFTER\
REM Output: C:\Temp\EQ_Snapshots\EQ_Diff_Report.txt
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0EQ-Diff.ps1"
pause
