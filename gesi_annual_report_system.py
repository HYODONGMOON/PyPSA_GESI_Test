# -*- coding: utf-8 -*-
"""
GESI 연구소 Annual Report 데이터베이스 시스템
https://gesi.kr/ 홈페이지와 연동하여 신규 콘텐츠를 자동 수집 및 저장

주요 기능:
1. Library (보고서) 수집
2. Projects (프로젝트/과제) 수집  
3. Events (행사) 수집
4. Annual Report 생성 및 Excel 내보내기
"""

import sqlite3
import pandas as pd
import hashlib
import json
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional
import time
import re

# Selenium 사용 (동적 콘텐츠 로딩)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException


class GESIDatabase:
    """GESI 데이터베이스 관리 클래스"""
    
    def __init__(self, db_path: str = "gesi_annual_report.db"):
        self.db_path = db_path
        self.conn = None
        self.cursor = None
        self._connect()
        self._create_tables()
    
    def _connect(self):
        """데이터베이스 연결"""
        self.conn = sqlite3.connect(self.db_path)
        self.cursor = self.conn.cursor()
    
    def _create_tables(self):
        """데이터베이스 테이블 생성"""
        
        # 보고서/발간물 테이블 (Library)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS library (
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
            )
        """)
        
        # 프로젝트/과제 테이블
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS projects (
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
            )
        """)
        
        # 행사/이벤트 테이블
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS events (
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
            )
        """)
        
        # 수집 이력 테이블
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS collection_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                collection_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                content_type TEXT NOT NULL,
                items_collected INTEGER,
                new_items INTEGER,
                status TEXT,
                notes TEXT
            )
        """)
        
        self.conn.commit()
    
    def _generate_hash(self, content: Dict) -> str:
        """콘텐츠 고유 해시 생성"""
        hash_string = json.dumps(content, sort_keys=True, ensure_ascii=False)
        return hashlib.md5(hash_string.encode()).hexdigest()
    
    def add_library_item(self, data: Dict) -> bool:
        """보고서 추가"""
        content_hash = self._generate_hash({
            'title': data.get('title'),
            'published_date': data.get('published_date')
        })
        
        try:
            self.cursor.execute("""
                INSERT INTO library (
                    content_hash, no, category, title, author,
                    published_date, views, url, file_url, summary
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                content_hash,
                data.get('no'),
                data.get('category'),
                data.get('title'),
                data.get('author'),
                data.get('published_date'),
                data.get('views'),
                data.get('url'),
                data.get('file_url'),
                data.get('summary')
            ))
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
    
    def add_project(self, data: Dict) -> bool:
        """프로젝트 추가"""
        content_hash = self._generate_hash({
            'project_name': data.get('project_name'),
            'year': data.get('year')
        })
        
        try:
            self.cursor.execute("""
                INSERT INTO projects (
                    content_hash, project_name, project_code, year,
                    start_date, end_date, status, project_type,
                    funding_agency, principal_investigator,
                    description, budget, url
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                content_hash,
                data.get('project_name'),
                data.get('project_code'),
                data.get('year'),
                data.get('start_date'),
                data.get('end_date'),
                data.get('status'),
                data.get('project_type'),
                data.get('funding_agency'),
                data.get('principal_investigator'),
                data.get('description'),
                data.get('budget'),
                data.get('url')
            ))
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
    
    def add_event(self, data: Dict) -> bool:
        """이벤트 추가"""
        content_hash = self._generate_hash({
            'event_name': data.get('event_name'),
            'event_date': data.get('event_date')
        })
        
        try:
            self.cursor.execute("""
                INSERT INTO events (
                    content_hash, event_name, event_type, event_date,
                    location, description, organizer, participants, url
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                content_hash,
                data.get('event_name'),
                data.get('event_type'),
                data.get('event_date'),
                data.get('location'),
                data.get('description'),
                data.get('organizer'),
                data.get('participants'),
                data.get('url')
            ))
            self.conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
    
    def log_collection(self, content_type: str, total: int, new: int, 
                       status: str = "success", notes: str = ""):
        """수집 이력 기록"""
        self.cursor.execute("""
            INSERT INTO collection_history (
                content_type, items_collected, new_items, status, notes
            ) VALUES (?, ?, ?, ?, ?)
        """, (content_type, total, new, status, notes))
        self.conn.commit()
    
    def get_library(self, year: int = None, category: str = None) -> pd.DataFrame:
        """보고서 조회"""
        query = "SELECT * FROM library WHERE 1=1"
        params = []
        
        if year:
            query += " AND strftime('%Y', published_date) = ?"
            params.append(str(year))
        
        if category:
            query += " AND category = ?"
            params.append(category)
        
        query += " ORDER BY published_date DESC"
        
        return pd.read_sql_query(query, self.conn, params=params)
    
    def get_projects(self, year: int = None, status: str = None) -> pd.DataFrame:
        """프로젝트 조회"""
        query = "SELECT * FROM projects WHERE 1=1"
        params = []
        
        if year:
            query += " AND year = ?"
            params.append(str(year))
        
        if status:
            query += " AND status = ?"
            params.append(status)
        
        query += " ORDER BY year DESC, project_name"
        
        return pd.read_sql_query(query, self.conn, params=params)
    
    def get_events(self, year: int = None) -> pd.DataFrame:
        """이벤트 조회"""
        query = "SELECT * FROM events WHERE 1=1"
        params = []
        
        if year:
            query += " AND strftime('%Y', event_date) = ?"
            params.append(str(year))
        
        query += " ORDER BY event_date DESC"
        
        return pd.read_sql_query(query, self.conn, params=params)
    
    def close(self):
        """데이터베이스 연결 종료"""
        if self.conn:
            self.conn.close()


class GESIScraper:
    """GESI 홈페이지 스크래퍼 (Selenium 사용)"""
    
    def __init__(self, headless: bool = True):
        """
        초기화
        
        Args:
            headless: 헤드리스 모드 사용 여부
        """
        self.base_url = "https://gesi.kr"
        self.driver = None
        self.headless = headless
        self._init_driver()
    
    def _init_driver(self):
        """Selenium WebDriver 초기화"""
        chrome_options = Options()
        
        if self.headless:
            chrome_options.add_argument('--headless')
        
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--window-size=1920,1080')
        chrome_options.add_argument('--lang=ko-KR')
        chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        
        try:
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.implicitly_wait(10)
        except Exception as e:
            print(f"❌ Chrome WebDriver 초기화 실패: {e}")
            print("💡 해결 방법:")
            print("   1. Chrome 브라우저가 설치되어 있는지 확인")
            print("   2. ChromeDriver가 설치되어 있는지 확인")
            print("   3. pip install selenium 실행")
            raise
    
    def scrape_library(self, max_pages: int = 5) -> List[Dict]:
        """
        Library 페이지에서 보고서 수집
        
        Args:
            max_pages: 수집할 최대 페이지 수
            
        Returns:
            보고서 정보 리스트
        """
        print("📚 Library (보고서) 수집 중...")
        url = f"{self.base_url}/library"
        
        all_items = []
        
        for page in range(1, max_pages + 1):
            try:
                page_url = f"{url}?page={page}"
                print(f"  페이지 {page} 로딩: {page_url}")
                
                self.driver.get(page_url)
                time.sleep(2)  # 동적 콘텐츠 로딩 대기
                
                # 테이블에서 데이터 추출
                try:
                    # tbody 내의 모든 행 찾기
                    rows = self.driver.find_elements(By.CSS_SELECTOR, "tbody tr")
                    
                    if not rows:
                        print(f"  페이지 {page}에 데이터 없음")
                        break
                    
                    for row in rows:
                        try:
                            cells = row.find_elements(By.TAG_NAME, "td")
                            
                            if len(cells) < 5:
                                continue
                            
                            # 데이터 추출
                            no = cells[0].text.strip()
                            category = cells[1].text.strip()
                            
                            # 제목과 URL
                            title_elem = cells[2].find_element(By.TAG_NAME, "a")
                            title = title_elem.text.strip()
                            item_url = title_elem.get_attribute("href")
                            
                            author = cells[3].text.strip()
                            pub_date = cells[4].text.strip()
                            views = cells[5].text.strip() if len(cells) > 5 else "0"
                            
                            # 날짜 형식 정규화 (YYYY-MM-DD)
                            pub_date = self._normalize_date(pub_date)
                            
                            item = {
                                'no': no,
                                'category': category,
                                'title': title,
                                'author': author,
                                'published_date': pub_date,
                                'views': self._extract_number(views),
                                'url': item_url,
                                'file_url': None,
                                'summary': None
                            }
                            
                            all_items.append(item)
                            
                        except Exception as e:
                            print(f"    ⚠ 행 파싱 오류: {e}")
                            continue
                    
                    print(f"  ✓ 페이지 {page}: {len(rows)}개 항목 수집")
                    
                except Exception as e:
                    print(f"  ⚠ 페이지 {page} 처리 오류: {e}")
                    break
                
                # 다음 페이지 확인
                try:
                    next_button = self.driver.find_element(By.CSS_SELECTOR, "a[class*='next']")
                    if not next_button.is_enabled():
                        break
                except:
                    break
                
            except Exception as e:
                print(f"  ⚠ 페이지 {page} 로딩 실패: {e}")
                break
        
        print(f"  총 {len(all_items)}개 보고서 수집 완료")
        return all_items
    
    def scrape_projects(self) -> List[Dict]:
        """
        Projects 페이지에서 프로젝트 수집
        
        Returns:
            프로젝트 정보 리스트
        """
        print("🔬 Projects (프로젝트) 수집 중...")
        url = f"{self.base_url}/projects"
        
        all_items = []
        
        try:
            self.driver.get(url)
            time.sleep(2)
            
            # 프로젝트 아이템 찾기
            items = self.driver.find_elements(By.CSS_SELECTOR, "div.item")
            
            for item in items:
                try:
                    # 제목
                    title_elem = item.find_element(By.CSS_SELECTOR, "h3, h4, .title, a")
                    title = title_elem.text.strip()
                    
                    # URL
                    try:
                        link_elem = item.find_element(By.TAG_NAME, "a")
                        item_url = link_elem.get_attribute("href")
                    except:
                        item_url = None
                    
                    # 설명
                    try:
                        desc_elem = item.find_element(By.CSS_SELECTOR, "p, .description")
                        description = desc_elem.text.strip()
                    except:
                        description = None
                    
                    # 연도 추출 (제목이나 설명에서)
                    year = self._extract_year(f"{title} {description or ''}")
                    
                    project = {
                        'project_name': title,
                        'project_code': None,
                        'year': year,
                        'start_date': None,
                        'end_date': None,
                        'status': '진행중',
                        'project_type': None,
                        'funding_agency': None,
                        'principal_investigator': None,
                        'description': description,
                        'budget': None,
                        'url': item_url
                    }
                    
                    all_items.append(project)
                    
                except Exception as e:
                    print(f"    ⚠ 프로젝트 파싱 오류: {e}")
                    continue
            
            print(f"  ✓ {len(all_items)}개 프로젝트 수집 완료")
            
        except Exception as e:
            print(f"  ⚠ 프로젝트 페이지 처리 오류: {e}")
        
        return all_items
    
    def scrape_events(self) -> List[Dict]:
        """
        Events 페이지에서 이벤트 수집
        
        Returns:
            이벤트 정보 리스트
        """
        print("📅 Events (행사) 수집 중...")
        url = f"{self.base_url}/EVENTS"  # 대문자 EVENTS
        
        all_items = []
        
        try:
            self.driver.get(url)
            time.sleep(2)
            
            # 이벤트 아이템 찾기 (구조에 따라 조정 필요)
            items = self.driver.find_elements(By.CSS_SELECTOR, "article, div.event-item, div.post")
            
            for item in items:
                try:
                    # 제목
                    title_elem = item.find_element(By.CSS_SELECTOR, "h3, h4, .title, a")
                    title = title_elem.text.strip()
                    
                    if not title:
                        continue
                    
                    # URL
                    try:
                        link_elem = item.find_element(By.TAG_NAME, "a")
                        item_url = link_elem.get_attribute("href")
                    except:
                        item_url = None
                    
                    # 날짜
                    try:
                        date_elem = item.find_element(By.CSS_SELECTOR, ".date, time")
                        event_date = date_elem.text.strip()
                        event_date = self._normalize_date(event_date)
                    except:
                        event_date = None
                    
                    # 설명
                    try:
                        desc_elem = item.find_element(By.CSS_SELECTOR, "p, .description")
                        description = desc_elem.text.strip()
                    except:
                        description = None
                    
                    event = {
                        'event_name': title,
                        'event_type': self._classify_event_type(title),
                        'event_date': event_date,
                        'location': None,
                        'description': description,
                        'organizer': 'GESI',
                        'participants': None,
                        'url': item_url
                    }
                    
                    all_items.append(event)
                    
                except Exception as e:
                    print(f"    ⚠ 이벤트 파싱 오류: {e}")
                    continue
            
            print(f"  ✓ {len(all_items)}개 이벤트 수집 완료")
            
        except Exception as e:
            print(f"  ⚠ 이벤트 페이지 처리 오류: {e}")
        
        return all_items
    
    def _normalize_date(self, date_str: str) -> Optional[str]:
        """날짜 문자열을 YYYY-MM-DD 형식으로 정규화"""
        if not date_str:
            return None
        
        # 공백 제거
        date_str = date_str.strip()
        
        # 이미 YYYY-MM-DD 형식
        if re.match(r'\d{4}-\d{2}-\d{2}', date_str):
            return date_str
        
        # YYYY.MM.DD 형식
        if re.match(r'\d{4}\.\d{2}\.\d{2}', date_str):
            return date_str.replace('.', '-')
        
        # YYYY/MM/DD 형식
        if re.match(r'\d{4}/\d{2}/\d{2}', date_str):
            return date_str.replace('/', '-')
        
        # YYYYMMDD 형식
        if re.match(r'\d{8}', date_str):
            return f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
        
        # 연도만 있는 경우
        year_match = re.search(r'20\d{2}', date_str)
        if year_match:
            return f"{year_match.group()}-01-01"
        
        return date_str
    
    def _extract_number(self, text: str) -> int:
        """텍스트에서 숫자 추출"""
        if not text:
            return 0
        
        numbers = re.findall(r'\d+', text)
        return int(numbers[0]) if numbers else 0
    
    def _extract_year(self, text: str) -> Optional[str]:
        """텍스트에서 연도 추출"""
        if not text:
            return None
        
        year_match = re.search(r'20\d{2}', text)
        return year_match.group() if year_match else None
    
    def _classify_event_type(self, title: str) -> str:
        """제목으로 이벤트 유형 분류"""
        if not title:
            return "기타"
        
        title_lower = title.lower()
        
        if '세미나' in title or 'seminar' in title_lower:
            return "세미나"
        elif '워크샵' in title or 'workshop' in title_lower:
            return "워크샵"
        elif '포럼' in title or 'forum' in title_lower:
            return "포럼"
        elif '컨퍼런스' in title or 'conference' in title_lower:
            return "컨퍼런스"
        elif '간담회' in title or '회의' in title:
            return "회의"
        elif '교육' in title or '강연' in title:
            return "교육/강연"
        else:
            return "기타"
    
    def close(self):
        """WebDriver 종료"""
        if self.driver:
            self.driver.quit()


class GESIAnnualReportCollector:
    """GESI Annual Report 수집 및 관리 클래스"""
    
    def __init__(self, headless: bool = True):
        """
        초기화
        
        Args:
            headless: 브라우저 헤드리스 모드 사용 여부
        """
        self.db = GESIDatabase()
        self.scraper = GESIScraper(headless=headless)
    
    def update_library(self, max_pages: int = 5) -> Dict:
        """보고서 업데이트"""
        print("\n" + "="*60)
        print("📚 Library (보고서) 업데이트")
        print("="*60)
        
        try:
            items = self.scraper.scrape_library(max_pages=max_pages)
            
            total = len(items)
            new_count = 0
            
            for item in items:
                if self.db.add_library_item(item):
                    new_count += 1
                    print(f"  ✓ 신규 보고서: {item['title']}")
            
            self.db.log_collection('library', total, new_count)
            
            return {'total': total, 'new': new_count}
            
        except Exception as e:
            print(f"❌ 보고서 수집 오류: {e}")
            self.db.log_collection('library', 0, 0, 'error', str(e))
            return {'total': 0, 'new': 0}
    
    def update_projects(self) -> Dict:
        """프로젝트 업데이트"""
        print("\n" + "="*60)
        print("🔬 Projects (프로젝트) 업데이트")
        print("="*60)
        
        try:
            items = self.scraper.scrape_projects()
            
            total = len(items)
            new_count = 0
            
            for item in items:
                if self.db.add_project(item):
                    new_count += 1
                    print(f"  ✓ 신규 프로젝트: {item['project_name']}")
            
            self.db.log_collection('projects', total, new_count)
            
            return {'total': total, 'new': new_count}
            
        except Exception as e:
            print(f"❌ 프로젝트 수집 오류: {e}")
            self.db.log_collection('projects', 0, 0, 'error', str(e))
            return {'total': 0, 'new': 0}
    
    def update_events(self) -> Dict:
        """이벤트 업데이트"""
        print("\n" + "="*60)
        print("📅 Events (행사) 업데이트")
        print("="*60)
        
        try:
            items = self.scraper.scrape_events()
            
            total = len(items)
            new_count = 0
            
            for item in items:
                if self.db.add_event(item):
                    new_count += 1
                    print(f"  ✓ 신규 이벤트: {item['event_name']}")
            
            self.db.log_collection('events', total, new_count)
            
            return {'total': total, 'new': new_count}
            
        except Exception as e:
            print(f"❌ 이벤트 수집 오류: {e}")
            self.db.log_collection('events', 0, 0, 'error', str(e))
            return {'total': 0, 'new': 0}
    
    def update_all(self, max_library_pages: int = 5) -> Dict:
        """모든 콘텐츠 업데이트"""
        print("\n" + "="*70)
        print("🔄 GESI Annual Report 데이터베이스 전체 업데이트")
        print(f"⏰ 수집 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("="*70)
        
        results = {
            'library': self.update_library(max_pages=max_library_pages),
            'projects': self.update_projects(),
            'events': self.update_events()
        }
        
        print("\n" + "="*70)
        print("✅ 업데이트 완료!")
        print(f"   📚 보고서: {results['library']['new']}/{results['library']['total']} 신규")
        print(f"   🔬 프로젝트: {results['projects']['new']}/{results['projects']['total']} 신규")
        print(f"   📅 이벤트: {results['events']['new']}/{results['events']['total']} 신규")
        print("="*70 + "\n")
        
        return results
    
    def generate_annual_report(self, year: int = None) -> Dict:
        """
        Annual Report 데이터 생성
        
        Args:
            year: 대상 연도 (None이면 현재 연도)
            
        Returns:
            연간 보고서 데이터
        """
        year = year or datetime.now().year
        
        print(f"\n📊 {year}년 Annual Report 데이터 생성 중...")
        
        # 데이터 조회
        library = self.db.get_library(year=year)
        projects = self.db.get_projects(year=year)
        events = self.db.get_events(year=year)
        
        # 카테고리별 보고서 수
        category_counts = library['category'].value_counts().to_dict() if not library.empty else {}
        
        data = {
            'year': year,
            'summary': {
                'total_publications': len(library),
                'total_projects': len(projects),
                'total_events': len(events),
                'publications_by_category': category_counts
            },
            'library': library.to_dict('records') if not library.empty else [],
            'projects': projects.to_dict('records') if not projects.empty else [],
            'events': events.to_dict('records') if not events.empty else []
        }
        
        print(f"✓ 발간물: {len(library)}건")
        print(f"✓ 프로젝트: {len(projects)}건")
        print(f"✓ 행사: {len(events)}건")
        
        return data
    
    def export_to_excel(self, year: int = None, output_dir: str = "annual_reports"):
        """
        Annual Report를 Excel 파일로 내보내기
        
        Args:
            year: 대상 연도
            output_dir: 출력 디렉토리
        """
        year = year or datetime.now().year
        
        # 출력 디렉토리 생성
        Path(output_dir).mkdir(exist_ok=True)
        
        output_path = Path(output_dir) / f"GESI_Annual_Report_{year}.xlsx"
        
        print(f"\n💾 Excel 파일 생성 중: {output_path}")
        
        # 데이터 생성
        data = self.generate_annual_report(year)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 1. 요약 시트
            summary_data = {
                '구분': ['발간물', '프로젝트', '행사'],
                '건수': [
                    data['summary']['total_publications'],
                    data['summary']['total_projects'],
                    data['summary']['total_events']
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='요약', index=False)
            
            # 2. 발간물 시트
            if data['library']:
                library_df = pd.DataFrame(data['library'])
                # 필요한 컬럼만 선택
                cols = ['no', 'category', 'title', 'author', 'published_date', 'views', 'url']
                library_df = library_df[[c for c in cols if c in library_df.columns]]
                library_df.to_excel(writer, sheet_name='발간물', index=False)
            
            # 3. 프로젝트 시트
            if data['projects']:
                projects_df = pd.DataFrame(data['projects'])
                projects_df.to_excel(writer, sheet_name='프로젝트', index=False)
            
            # 4. 행사 시트
            if data['events']:
                events_df = pd.DataFrame(data['events'])
                events_df.to_excel(writer, sheet_name='행사', index=False)
        
        print(f"✅ Excel 파일 저장 완료: {output_path}")
        
        return output_path
    
    def close(self):
        """리소스 정리"""
        self.scraper.close()
        self.db.close()


# ============================================
# 메인 실행 함수
# ============================================

def main():
    """메인 실행 함수"""
    
    print("""
╔═══════════════════════════════════════════════════════════════╗
║                                                               ║
║         GESI Annual Report 데이터베이스 시스템                 ║
║         녹색에너지전략연구소 (https://gesi.kr)                 ║
║                                                               ║
╚═══════════════════════════════════════════════════════════════╝
    """)
    
    # 수집기 초기화 (headless=False로 하면 브라우저 확인 가능)
    collector = GESIAnnualReportCollector(headless=True)
    
    try:
        # 1. 모든 데이터 업데이트 (Library는 최근 3페이지만)
        results = collector.update_all(max_library_pages=3)
        
        # 2. 현재 연도 Annual Report 생성
        current_year = datetime.now().year
        collector.export_to_excel(year=current_year)
        
        # 3. 이전 연도도 생성 (선택적)
        # collector.export_to_excel(year=current_year - 1)
        
        print("\n" + "="*70)
        print("🎉 모든 작업이 완료되었습니다!")
        print("="*70)
        
    except Exception as e:
        print(f"\n❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        # 리소스 정리
        collector.close()


if __name__ == "__main__":
    main()

