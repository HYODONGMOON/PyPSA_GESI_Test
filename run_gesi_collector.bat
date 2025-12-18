@echo off
chcp 65001 > nul
echo ╔═══════════════════════════════════════════════════════════════╗
echo ║                                                               ║
echo ║         GESI Annual Report 데이터 수집 시스템                  ║
echo ║                                                               ║
echo ╚═══════════════════════════════════════════════════════════════╝
echo.

echo [1/3] 필수 패키지 확인 중...
python -c "import selenium, pandas, openpyxl" 2>nul
if %errorlevel% neq 0 (
    echo ❌ 필수 패키지가 설치되지 않았습니다.
    echo 설치를 시작합니다...
    pip install -r requirements_annual_report.txt
)

echo.
echo [2/3] 데이터 수집 시작...
python gesi_annual_report_system.py

echo.
echo [3/3] 완료!
echo.
echo 생성된 파일:
echo   - annual_reports\GESI_Annual_Report_2024.xlsx
echo   - gesi_annual_report.db
echo.

pause

