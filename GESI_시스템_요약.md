# GESI Annual Report 데이터베이스 시스템 - 간단 요약

## 🎯 목적

GESI 홈페이지(https://gesi.kr)에서 신규 콘텐츠를 자동으로 수집하여 데이터베이스에 저장하고, Annual Report 작성을 위한 데이터를 제공하는 시스템입니다.

## 📁 핵심 파일 3개

| 파일명 | 용도 | 실행 방법 |
|--------|------|-----------|
| **gesi_annual_report_system.py** | 메인 시스템 | `python gesi_annual_report_system.py` |
| **demo_gesi_system.py** | 기능 테스트/데모 | `python demo_gesi_system.py` |
| **run_gesi_collector.bat** | 원클릭 실행 (Windows) | 더블클릭 |

## ⚡ 10초 시작 가이드

```bash
# 1. 패키지 설치
pip install selenium pandas openpyxl beautifulsoup4 requests

# 2. 실행
python gesi_annual_report_system.py

# 3. 결과 확인
# → gesi_annual_report.db (데이터베이스)
# → annual_reports/GESI_Annual_Report_2024.xlsx (엑셀 파일)
```

## 🔄 동작 방식

```
1. GESI 홈페이지 접속
   ↓
2. Library/Projects/Events 페이지에서 데이터 수집
   ↓
3. 중복 확인 후 신규 데이터만 DB 저장
   ↓
4. 연도별 데이터 집계
   ↓
5. Excel 파일로 내보내기
```

## 📊 수집 데이터

### Library (보고서)
- ✅ 제목, 저자, 발간일
- ✅ 카테고리, 조회수
- ✅ URL

### Projects (프로젝트)
- ✅ 프로젝트명, 연도
- ✅ 상태 (진행중/완료)
- ✅ 설명

### Events (행사)
- ✅ 행사명, 날짜
- ✅ 행사 유형
- ✅ 설명

## 💻 주요 기능 코드

### 전체 데이터 수집
```python
from gesi_annual_report_system import GESIAnnualReportCollector

collector = GESIAnnualReportCollector(headless=True)
try:
    collector.update_all(max_library_pages=3)
finally:
    collector.close()
```

### Excel 파일 생성
```python
collector = GESIAnnualReportCollector(headless=True)
try:
    collector.export_to_excel(year=2024)
finally:
    collector.close()
```

### 데이터 조회
```python
from gesi_annual_report_system import GESIDatabase

db = GESIDatabase()
try:
    reports = db.get_library(year=2024)
    print(f"2024년 보고서: {len(reports)}건")
finally:
    db.close()
```

## 🎨 Excel 출력 형식

생성되는 `GESI_Annual_Report_2024.xlsx` 파일:

| 시트명 | 내용 |
|--------|------|
| **요약** | 발간물/프로젝트/행사 건수 통계 |
| **발간물** | 번호, 카테고리, 제목, 저자, 날짜, 조회수, URL |
| **프로젝트** | 프로젝트명, 연도, 상태, 설명 등 |
| **행사** | 행사명, 유형, 날짜, 장소, 설명 등 |

## ⏰ 자동화 설정

### Windows 작업 스케줄러 (매일 자동 실행)

1. `Win + R` → `taskschd.msc` 입력
2. "작업 만들기" 클릭
3. **트리거**: 매일 오전 9시
4. **동작**: `run_gesi_collector.bat` 실행
5. 완료!

## 🐛 문제 해결

| 문제 | 해결 방법 |
|------|-----------|
| ChromeDriver 오류 | `pip install webdriver-manager` |
| 데이터 수집 안됨 | `headless=False`로 변경하여 브라우저 확인 |
| Excel 오류 | `pip install openpyxl` |
| 패키지 오류 | `pip install -r requirements_annual_report.txt` |

## 📚 상세 문서

- **README_GESI_Annual_Report.md**: 전체 시스템 개요
- **GESI_Annual_Report_사용가이드.md**: 상세 사용 방법
- **demo_gesi_system.py**: 대화형 데모

## 🔑 핵심 장점

✅ **자동화**: 한 번 설정하면 자동으로 데이터 수집  
✅ **중복 방지**: 똑같은 데이터를 여러 번 수집해도 한 번만 저장  
✅ **쉬운 사용**: Python 한 줄로 실행  
✅ **Excel 연동**: Annual Report 작성 시 바로 사용 가능  
✅ **확장 가능**: 새로운 데이터 타입 추가 가능

## 📞 도움이 필요하면

1. `demo_gesi_system.py` 실행 → 대화형 데모로 기능 테스트
2. `GESI_Annual_Report_사용가이드.md` 참고 → 상세 설명
3. 문의: 시스템 관리자

---

**"신규 보고서가 5개 올라오면, 코드 실행으로 자동으로 DB에 추가!"**

