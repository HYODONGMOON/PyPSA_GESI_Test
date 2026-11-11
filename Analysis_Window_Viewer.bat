@echo off
rem Analysis Window viewer. Drag a result folder onto this file to open it directly.
cd /d "%~dp0"
"C:\ProgramData\anaconda3\envs\pypsa_env\python.exe" "%~dp0analysis_window_viewer.py" %*
if errorlevel 1 pause
