@echo off
REM Merges the handles.csv you downloaded from the helper page into players.xlsx.
cd /d "%~dp0"
python apply_handles.py
pause
