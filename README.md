# GESI Annual Report 데이터베이스 시스템

**녹색에너지전략연구소(GESI)** 홈페이지와 연동하여 연간 보고서 작성을 위한 데이터를 자동으로 수집하고 관리하는 시스템입니다.

## ✨ 주요 특징

- 🔄 **자동 데이터 수집**: GESI 홈페이지에서 신규 콘텐츠 자동 수집
- 🚫 **중복 방지**: 해시 기반 중복 검사로 신규 항목만 추가
- 📊 **데이터베이스 관리**: SQLite 기반 체계적인 데이터 관리
- 📈 **Annual Report 생성**: 연도별 데이터 집계 및 Excel 파일 자동 생성
- ⏰ **정기 실행 지원**: 스케줄러 연동으로 자동화 가능

## 📦 파일 구성

```
├── gesi_annual_report_system.py      # 메인 시스템 코드
├── demo_gesi_system.py               # 데모/테스트 스크립트
├── analyze_gesi_website.py           # 웹사이트 구조 분석 도구
├── requirements_annual_report.txt    # 필수 패키지 목록
├── run_gesi_collector.bat            # Windows 실행 배치 파일
├── GESI_Annual_Report_사용가이드.md  # 상세 사용 가이드
└── README_GESI_Annual_Report.md      # 이 파일
```

## 🚀 빠른 시작

### 1단계: 패키지 설치

```bash
pip install -r requirements_annual_report.txt
```

### 2단계: 실행

**방법 1: 배치 파일 실행 (Windows)**
```bash
run_gesi_collector.bat
```

**방법 2: Python 직접 실행**
```bash
python gesi_annual_report_system.py
```

**방법 3: 데모 실행 (기능 테스트)**
```bash
python demo_gesi_system.py
```

### 3단계: 결과 확인

- **데이터베이스**: `gesi_annual_report.db`
- **Excel 파일**: `annual_reports/GESI_Annual_Report_2024.xlsx`

## 💡 사용 예시

### 예시 1: 전체 데이터 수집

```python
from gesi_annual_report_system import GESIAnnualReportCollector

collector = GESIAnnualReportCollector(headless=True)

try:
    # 모든 데이터 수집 (Library 5페이지까지)
    results = collector.update_all(max_library_pages=5)
    
    # 결과 확인
    print(f"보고서: {results['library']['new']}/{results['library']['total']} 신규")
    print(f"프로젝트: {results['projects']['new']}/{results['projects']['total']} 신규")
    print(f"이벤트: {results['events']['new']}/{results['events']['total']} 신규")
    
finally:
    collector.close()
```

### 예시 2: 특정 연도 Annual Report 생성

```python
from gesi_annual_report_system import GESIAnnualReportCollector

collector = GESIAnnualReportCollector(headless=True)

try:
    # 2024년 보고서 생성
    collector.export_to_excel(year=2024, output_dir="reports_2024")
    
    # 2023년 보고서 생성
    collector.export_to_excel(year=2023, output_dir="reports_2023")
    
finally:
    collector.close()
```

### 예시 3: 데이터 조회

```python
from gesi_annual_report_system import GESIDatabase

db = GESIDatabase()

try:
    # 2024년 보고서 조회
    reports_2024 = db.get_library(year=2024)
    print(f"2024년 발간물: {len(reports_2024)}건")
    
    # 카테고리별 조회
    research_reports = db.get_library(category="연구보고서")
    print(f"연구보고서: {len(research_reports)}건")
    
    # 진행중인 프로젝트 조회
    ongoing_projects = db.get_projects(status="진행중")
    print(f"진행중인 프로젝트: {len(ongoing_projects)}건")
    
finally:
    db.close()
```

## 📊 수집되는 데이터

### 1. Library (보고서/발간물)
- 보고서 번호, 카테고리, 제목
- 저자, 발간일, 조회수
- 상세 페이지 URL

### 2. Projects (프로젝트/과제)
- 프로젝트명, 과제 코드
- 연도, 상태 (진행중/완료)
- 설명, 예산, 연구자 정보

### 3. Events (행사/이벤트)
- 행사명, 행사 유형
- 날짜, 장소
- 설명, 주최자

## 🔧 시스템 구조

```
┌─────────────────────────────────────────────────┐
│           GESI Annual Report System             │
└─────────────────────────────────────────────────┘
                        │
        ┌───────────────┼───────────────┐
        │               │               │
┌───────▼──────┐ ┌─────▼──────┐ ┌─────▼──────┐
│  Web Scraper │ │  Database  │ │   Excel    │
│  (Selenium)  │ │  (SQLite)  │ │  Export    │
└──────────────┘ └────────────┘ └────────────┘
        │               │               │
┌───────▼──────────────┐│               │
│  https://gesi.kr     ││               │
│  ├─ /library         ││               │
│  ├─ /projects        ││               │
│  └─ /EVENTS          ││               │
└──────────────────────┘│               │
                        │               │
                ┌───────▼───────┐       │
                │  .db 파일     │       │
                │  ├─ library   │       │
                │  ├─ projects  │       │
                │  ├─ events    │       │
                │  └─ history   │       │
                └───────────────┘       │
                                        │
                                ┌───────▼────────┐
                                │  .xlsx 파일    │
                                │  ├─ 요약       │
                                │  ├─ 발간물     │
                                │  ├─ 프로젝트   │
                                │  └─ 행사       │
                                └────────────────┘
```

## 🔄 자동화 설정

### Windows 작업 스케줄러

1. **작업 스케줄러 실행**: `taskschd.msc`
2. **작업 만들기**:
   - 이름: "GESI 데이터 수집"
   - 트리거: 매일 오전 9시
   - 동작: `run_gesi_collector.bat` 실행
   - 시작 위치: 프로젝트 폴더

### Python 스케줄러 (schedule 라이브러리)

```python
import schedule
import time

def collect_data():
    from gesi_annual_report_system import GESIAnnualReportCollector
    collector = GESIAnnualReportCollector(headless=True)
    try:
        collector.update_all(max_library_pages=3)
    finally:
        collector.close()

# 매일 오전 9시 실행
schedule.every().day.at("09:00").do(collect_data)

# 매주 월요일 실행
schedule.every().monday.at("09:00").do(collect_data)

while True:
    schedule.run_pending()
    time.sleep(60)
```

## 📋 데이터베이스 스키마

### library (보고서)
```sql
CREATE TABLE library (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    content_hash TEXT UNIQUE NOT NULL,
    no TEXT,
    category TEXT,
    title TEXT NOT NULL,
    author TEXT,
    published_date DATE,
    views INTEGER,
    url TEXT,
    file_url TEXT,
    summary TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

### projects (프로젝트)
```sql
CREATE TABLE projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    content_hash TEXT UNIQUE NOT NULL,
    project_name TEXT NOT NULL,
    project_code TEXT,
    year TEXT,
    start_date DATE,
    end_date DATE,
    status TEXT,
    project_type TEXT,
    funding_agency TEXT,
    principal_investigator TEXT,
    description TEXT,
    budget TEXT,
    url TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

### events (행사)
```sql
CREATE TABLE events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    content_hash TEXT UNIQUE NOT NULL,
    event_name TEXT NOT NULL,
    event_type TEXT,
    event_date DATE,
    location TEXT,
    description TEXT,
    organizer TEXT,
    participants TEXT,
    url TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

## ⚠️ 문제 해결

### ChromeDriver 오류
```
❌ Chrome WebDriver 초기화 실패
```

**해결:**
```bash
pip install webdriver-manager
```

### Selenium이 데이터를 찾지 못함

**해결:**
1. `headless=False`로 브라우저 확인
2. GESI 웹사이트 구조 변경 가능성 확인
3. `analyze_gesi_website.py` 실행하여 구조 재분석

### 패키지 설치 오류

**해결:**
```bash
pip install --upgrade pip
pip install -r requirements_annual_report.txt --upgrade
```

## 📚 참고 문서

- **상세 사용 가이드**: `GESI_Annual_Report_사용가이드.md`
- **웹사이트 분석 도구**: `analyze_gesi_website.py`
- **데모 스크립트**: `demo_gesi_system.py`

## 🔐 보안 및 윤리

- 이 시스템은 **공개된 웹사이트**에서만 데이터를 수집합니다
- **적절한 딜레이**(2초)를 두어 서버 부하를 최소화합니다
- 개인정보는 수집하지 않습니다
- 수집된 데이터는 연구소 내부 용도로만 사용됩니다

## 🛠️ 기술 스택

- **Python 3.8+**
- **Selenium**: 동적 웹 페이지 스크래핑
- **SQLite**: 데이터베이스
- **Pandas**: 데이터 처리 및 분석
- **openpyxl**: Excel 파일 생성

## 📝 라이센스

이 코드는 GESI 연구소 내부 사용을 위해 작성되었습니다.

## 📞 문의

시스템 관련 문의:
- **개발자**: [담당자 이름]
- **이메일**: [이메일 주소]
- **연구소**: 녹색에너지전략연구소 (https://gesi.kr)

---

**버전**: 1.0.0  
**최종 업데이트**: 2024-12-18  
**개발**: Python 3.x

