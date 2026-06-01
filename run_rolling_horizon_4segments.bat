@echo off
REM PyPSA GESI Rolling Horizon 분석 실행 (4개 구간, CPLEX 사용)
REM 
REM 이 스크립트는 conda 환경에서 Rolling Horizon 분석을 실행합니다.
REM 4개 구간으로 나누어 분석하면 각 구간당 약 2190시간(약 3개월)입니다.

echo ======================================================================
echo PyPSA GESI Rolling Horizon 분석 (CPLEX) - 4개 구간
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
echo [옵션] --rolling-horizon --segments 4
echo [정보] 각 구간: 약 2190시간 (약 3개월)
echo.

REM Rolling Horizon 분석 실행 (4개 구간)
"%CONDA_PYTHON%" PyPSA_GUI.py --rolling-horizon --segments 4

echo.
echo ======================================================================
echo 분석 완료!
echo ======================================================================
pause
