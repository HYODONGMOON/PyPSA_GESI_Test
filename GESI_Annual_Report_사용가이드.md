# GESI Annual Report 데이터베이스 시스템 사용 가이드

## 📋 개요

GESI(녹색에너지전략연구소) 홈페이지(https://gesi.kr)와 연동하여 신규 콘텐츠를 자동으로 수집하고 데이터베이스에 저장하는 시스템입니다.

### 주요 기능

1. **자동 데이터 수집**
   - Library (보고서/발간물) 수집
   - Projects (프로젝트/과제) 수집
   - Events (행사/이벤트) 수집

2. **중복 방지**
   - 콘텐츠 해시를 사용하여 중복 데이터 자동 필터링
   - 신규 콘텐츠만 데이터베이스에 추가

3. **Annual Report 생성**
   - 연도별 데이터 조회 및 집계
   - Excel 파일로 자동 내보내기

## 🚀 설치 방법

### 1. 필수 패키지 설치

```bash
pip install -r requirements_annual_report.txt
```

또는 개별 설치:

```bash
pip install selenium beautifulsoup4 requests pandas openpyxl
```

### 2. ChromeDriver 설치

Selenium은 Chrome 브라우저를 제어하므로 ChromeDriver가 필요합니다.

**자동 설치 (권장):**
```bash
pip install webdriver-manager
```

**수동 설치:**
1. Chrome 브라우저 버전 확인 (chrome://version)
2. https://chromedriver.chromium.org/downloads 에서 동일 버전 다운로드
3. PATH에 추가 또는 프로젝트 폴더에 배치

## 💻 사용 방법

### 기본 실행

```python
python gesi_annual_report_system.py
```

실행하면 다음 작업이 자동으로 수행됩니다:
1. GESI 홈페이지에서 최신 데이터 수집
2. 데이터베이스에 신규 항목 저장
3. 현재 연도 Annual Report Excel 파일 생성

### 커스텀 실행 예시

```python
from gesi_annual_report_system import GESIAnnualReportCollector

# 초기화 (headless=False로 하면 브라우저가 보임)
collector = GESIAnnualReportCollector(headless=True)

try:
    # 1. 특정 콘텐츠만 업데이트
    collector.update_library(max_pages=5)  # Library만 업데이트 (5페이지)
    collector.update_projects()  # Projects만 업데이트
    collector.update_events()  # Events만 업데이트
    
    # 2. 특정 연도 Annual Report 생성
    collector.export_to_excel(year=2024, output_dir="reports_2024")
    collector.export_to_excel(year=2023, output_dir="reports_2023")
    
    # 3. 데이터 조회
    library_2024 = collector.db.get_library(year=2024)
    print(f"2024년 발간물: {len(library_2024)}건")
    
    projects_ongoing = collector.db.get_projects(status="진행중")
    print(f"진행중인 프로젝트: {len(projects_ongoing)}건")
    
finally:
    collector.close()
```

## 📊 데이터베이스 구조

### 테이블 설명

#### 1. `library` (보고서/발간물)
| 컬럼명 | 타입 | 설명 |
|--------|------|------|
| id | INTEGER | 고유 ID (자동 증가) |
| content_hash | TEXT | 중복 방지용 해시 |
| no | TEXT | 보고서 번호 |
| category | TEXT | 카테고리 (연구보고서, 정책보고서 등) |
| title | TEXT | 제목 |
| author | TEXT | 저자 |
| published_date | DATE | 발간일 |
| views | INTEGER | 조회수 |
| url | TEXT | 상세 페이지 URL |
| file_url | TEXT | 파일 다운로드 URL |
| summary | TEXT | 요약 |
| created_at | TIMESTAMP | 데이터베이스 등록일 |

#### 2. `projects` (프로젝트/과제)
| 컬럼명 | 타입 | 설명 |
|--------|------|------|
| id | INTEGER | 고유 ID |
| content_hash | TEXT | 중복 방지용 해시 |
| project_name | TEXT | 프로젝트명 |
| project_code | TEXT | 과제 코드 |
| year | TEXT | 연도 |
| start_date | DATE | 시작일 |
| end_date | DATE | 종료일 |
| status | TEXT | 상태 (진행중, 완료 등) |
| project_type | TEXT | 프로젝트 유형 |
| funding_agency | TEXT | 지원 기관 |
| principal_investigator | TEXT | 책임 연구자 |
| description | TEXT | 설명 |
| budget | TEXT | 예산 |
| url | TEXT | 상세 페이지 URL |

#### 3. `events` (행사/이벤트)
| 컬럼명 | 타입 | 설명 |
|--------|------|------|
| id | INTEGER | 고유 ID |
| content_hash | TEXT | 중복 방지용 해시 |
| event_name | TEXT | 행사명 |
| event_type | TEXT | 행사 유형 (세미나, 워크샵 등) |
| event_date | DATE | 행사 날짜 |
| location | TEXT | 장소 |
| description | TEXT | 설명 |
| organizer | TEXT | 주최자 |
| participants | TEXT | 참가자 |
| url | TEXT | 상세 페이지 URL |

#### 4. `collection_history` (수집 이력)
| 컬럼명 | 타입 | 설명 |
|--------|------|------|
| id | INTEGER | 고유 ID |
| collection_date | TIMESTAMP | 수집 일시 |
| content_type | TEXT | 콘텐츠 유형 |
| items_collected | INTEGER | 수집한 총 항목 수 |
| new_items | INTEGER | 신규 항목 수 |
| status | TEXT | 상태 (success, error) |
| notes | TEXT | 비고 |

## 🔄 정기 실행 설정

### Windows 작업 스케줄러

1. 작업 스케줄러 실행 (taskschd.msc)
2. "작업 만들기" 클릭
3. 트리거 설정 (예: 매일 오전 9시)
4. 동작 설정:
   - 프로그램: `python.exe` 경로
   - 인수: `gesi_annual_report_system.py` 경로
   - 시작 위치: 프로젝트 폴더 경로

### 배치 파일 사용

`run_gesi_collector.bat` 파일 생성:

```batch
@echo off
cd /d "C:\Users\Hyodong.Moon\Desktop\HDMOON\python workplace\PyPSA_GESI_Test"
python gesi_annual_report_system.py
pause
```

## 📈 출력 파일

### Excel 파일 구조

`GESI_Annual_Report_YYYY.xlsx` 파일에 다음 시트가 생성됩니다:

1. **요약 시트**
   - 발간물, 프로젝트, 행사 건수 요약

2. **발간물 시트**
   - 번호, 카테고리, 제목, 저자, 발간일, 조회수, URL

3. **프로젝트 시트**
   - 프로젝트명, 연도, 상태, 설명 등

4. **행사 시트**
   - 행사명, 유형, 날짜, 장소, 설명 등

## 🔧 커스터마이징

### 수집 페이지 수 조정

```python
# Library 10페이지까지 수집
collector.update_library(max_pages=10)
```

### 헤드리스 모드 비활성화 (브라우저 확인)

```python
# 브라우저 창이 보이도록 설정
collector = GESIAnnualReportCollector(headless=False)
```

### 특정 연도만 조회

```python
# 2023년 데이터만 조회
library_2023 = collector.db.get_library(year=2023)
projects_2023 = collector.db.get_projects(year=2023)
events_2023 = collector.db.get_events(year=2023)
```

### 카테고리별 필터링

```python
# 연구보고서만 조회
research_reports = collector.db.get_library(category="연구보고서")
```

## ⚠️ 주의사항

1. **웹사이트 구조 변경**
   - GESI 홈페이지 구조가 변경되면 스크래퍼 코드 수정 필요
   - `GESIScraper` 클래스의 CSS 선택자 업데이트

2. **과도한 요청 방지**
   - 적절한 딜레이 설정 (현재 2초)
   - 너무 많은 페이지 수집 시 서버 부하 고려

3. **ChromeDriver 버전**
   - Chrome 브라우저 업데이트 시 ChromeDriver도 업데이트 필요

4. **네트워크 연결**
   - 안정적인 인터넷 연결 필요

## 🐛 문제 해결

### Chrome WebDriver 오류

```
❌ Chrome WebDriver 초기화 실패
```

**해결 방법:**
1. Chrome 브라우저 설치 확인
2. ChromeDriver 설치: `pip install webdriver-manager`
3. 또는 수동으로 ChromeDriver 다운로드

### 데이터가 수집되지 않음

**확인 사항:**
1. 인터넷 연결 확인
2. GESI 홈페이지 접속 가능 여부 확인
3. `headless=False`로 브라우저 확인
4. CSS 선택자가 현재 웹사이트 구조와 일치하는지 확인

### Excel 파일 생성 오류

```
pip install openpyxl
```

## 📞 문의

시스템 관련 문의나 개선 사항이 있으시면 연락주세요.

---

**버전:** 1.0
**최종 수정일:** 2024-12-18

