@echo off
chcp 65001 > nul
echo.
echo ============================================================
echo   PyPSA Analysis 실행파일 빌드
echo   결과물: dist\PyPSA_Analysis\PyPSA_Analysis.exe
echo ============================================================
echo.

REM ── conda 환경 활성화 ─────────────────────────────────────────
call conda activate pypsa_env 2>nul
if errorlevel 1 (
    echo [오류] conda 환경 'pypsa_env' 활성화 실패.
    echo        Anaconda 가 설치되어 있고, pypsa_env 환경이 존재하는지 확인하세요.
    pause
    exit /b 1
)
echo [OK] conda 환경 pypsa_env 활성화 완료

REM ── PyInstaller 설치 확인 및 자동 설치 ───────────────────────
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [INFO] PyInstaller 미설치 - 지금 설치합니다...
    pip install pyinstaller
    if errorlevel 1 (
        echo [오류] PyInstaller 설치 실패. 인터넷 연결을 확인하세요.
        pause
        exit /b 1
    )
)
echo [OK] PyInstaller 준비 완료

REM ── 빌드 실행 ─────────────────────────────────────────────────
echo.
echo [빌드 시작] 패키지 크기가 크므로 수 분 이상 소요될 수 있습니다...
echo.

pyinstaller PyPSA_Analysis.spec --clean --noconfirm

if errorlevel 1 (
    echo.
    echo [오류] 빌드 실패. 위 오류 메시지를 확인하세요.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   빌드 완료!
echo.
echo   실행파일 위치: dist\PyPSA_Analysis\PyPSA_Analysis.exe
echo.
echo   [배포 방법]
echo   dist\PyPSA_Analysis\ 폴더 전체를 전달하세요.
echo   폴더 안에 interface.xlsx 가 포함되어 있습니다.
echo   사용자는 interface.xlsx 를 채운 뒤
echo   PyPSA_Analysis.exe 를 더블클릭하면 분석이 시작됩니다.
echo ============================================================
echo.
pause
