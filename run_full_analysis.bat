@echo off
REM PyPSA GESI 전체 분석 실행 (CPLEX 사용)
REM 
REM 이 스크립트는 conda 환경에서 전체 분석을 실행합니다.

echo ======================================================================
echo PyPSA GESI 전체 분석 (CPLEX)
echo ======================================================================
echo.

REM conda 환경의 Python 경로
set CONDA_PYTHON=C:\ProgramData\anaconda3\envs\pypsa_env\python.exe

REM 파일 존재 확인
if not exist "%CONDA_PYTHON%" (
    echo [오류] conda 환경을 찾을 수 없습니다: %CONDA_PYTHON%
    echo        conda 환경을 설치하거나 경로를 수정하세요.
    pause
    exit /b 1
)

echo [실행] %CONDA_PYTHON%
echo.

REM 전체 분석 실행
"%CONDA_PYTHON%" PyPSA_GUI.py

echo.
echo ======================================================================
echo 분석 완료!
echo ======================================================================
pause
