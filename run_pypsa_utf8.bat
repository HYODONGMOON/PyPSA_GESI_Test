@echo off
REM Set UTF-8 encoding for Python output
chcp 65001
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

echo ====================================
echo PyPSA GESI Analysis (UTF-8 Mode)
echo ====================================
echo.

python PyPSA_GUI.py

pause
