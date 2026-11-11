@echo off
rem Analysis Window viewer.
rem  - In the project folder: opens the last used result folder.
rem  - Copied into a result folder: opens that folder.
rem  - Drag a result folder onto this file to open it.
set "PY=C:\ProgramData\anaconda3\envs\pypsa_env\python.exe"
set "VIEWER=%~dp0analysis_window_viewer.py"
set "TARGET=%~1"
if not exist "%VIEWER%" (
    set "VIEWER=C:\Users\Hyodong.Moon\Desktop\HDMOON\python workplace\PyPSA_GESI_Test\analysis_window_viewer.py"
    if "%~1"=="" set "TARGET=%~dp0."
)
for %%I in ("%VIEWER%") do set "PYTHONPATH=%%~dpI"
if "%TARGET%"=="" (
    "%PY%" "%VIEWER%"
) else (
    "%PY%" "%VIEWER%" "%TARGET%"
)
if errorlevel 1 pause
