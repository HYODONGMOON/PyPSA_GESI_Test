# -*- coding: utf-8 -*-
"""
GESI 홈페이지 구조 분석 스크립트
https://gesi.kr/ 사이트의 구조를 파악하여 스크래핑 전략 수립
"""

import requests
from bs4 import BeautifulSoup
import json
from urllib.parse import urljoin, urlparse
import time


def analyze_page_structure(url: str, name: str = ""):
    """페이지 구조 분석"""
    print(f"\n{'='*70}")
    print(f"📄 분석 중: {name or url}")
    print(f"{'='*70}")
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        print(f"✓ 페이지 로드 성공 (상태 코드: {response.status_code})")
        print(f"✓ 인코딩: {response.encoding}")
        
        # 1. 메인 컨테이너 찾기
        print("\n[1] 주요 컨테이너 구조:")
        containers = [
            'main', 'article', 'section', 
            'div[class*="content"]', 'div[class*="container"]',
            'div[class*="wrapper"]', 'div[id*="content"]'
        ]
        
        for selector in containers:
            elements = soup.select(selector)
            if elements:
                print(f"  - {selector}: {len(elements)}개 발견")
                if len(elements) <= 3:
                    for elem in elements:
                        classes = elem.get('class', [])
                        id_attr = elem.get('id', '')
                        print(f"    • class={classes}, id={id_attr}")
        
        # 2. 리스트 아이템 찾기 (보고서, 프로젝트, 이벤트 목록)
        print("\n[2] 목록 아이템 패턴:")
        list_patterns = [
            'article', 'li', 
            'div[class*="item"]', 'div[class*="card"]',
            'div[class*="post"]', 'div[class*="list"]',
            'tr', 'tbody tr'
        ]
        
        for pattern in list_patterns:
            elements = soup.select(pattern)
            if elements and len(elements) > 1:
                print(f"  - {pattern}: {len(elements)}개")
                if len(elements) <= 20:  # 너무 많지 않으면 첫 항목 분석
                    first = elements[0]
                    classes = first.get('class', [])
                    print(f"    첫 항목 class: {classes}")
                    
                    # 제목 찾기
                    title_elem = (first.find('h1') or first.find('h2') or 
                                 first.find('h3') or first.find('h4') or
                                 first.find('a'))
                    if title_elem:
                        title_text = title_elem.get_text(strip=True)[:50]
                        print(f"    제목 샘플: {title_text}...")
        
        # 3. 링크 분석
        print("\n[3] 주요 링크:")
        links = soup.find_all('a', href=True)
        link_categories = {
            'library': [],
            'project': [],
            'research': [],
            'publication': [],
            'report': [],
            'event': [],
            'activity': [],
            'news': []
        }
        
        for link in links:
            href = link.get('href', '')
            text = link.get_text(strip=True)
            
            # URL 정규화
            full_url = urljoin(url, href)
            
            # 카테고리별 분류
            href_lower = href.lower()
            for category in link_categories.keys():
                if category in href_lower or category in text.lower():
                    if full_url not in link_categories[category]:
                        link_categories[category].append({
                            'url': full_url,
                            'text': text[:50]
                        })
        
        for category, items in link_categories.items():
            if items:
                print(f"  [{category}] {len(items)}개 링크:")
                for item in items[:3]:  # 처음 3개만 표시
                    print(f"    • {item['text']}: {item['url']}")
        
        # 4. 테이블 구조 (보고서가 테이블로 표시될 수 있음)
        print("\n[4] 테이블 구조:")
        tables = soup.find_all('table')
        for i, table in enumerate(tables[:3]):  # 최대 3개만
            rows = table.find_all('tr')
            print(f"  테이블 {i+1}: {len(rows)}개 행")
            if rows:
                headers = [th.get_text(strip=True) for th in rows[0].find_all(['th', 'td'])]
                if headers:
                    print(f"    헤더: {headers}")
        
        # 5. 페이지네이션
        print("\n[5] 페이지네이션:")
        pagination_patterns = [
            'nav[class*="pag"]', 'div[class*="pag"]',
            'ul[class*="pag"]', 'a[class*="next"]',
            'a[class*="prev"]', 'a[class*="page"]'
        ]
        
        for pattern in pagination_patterns:
            elements = soup.select(pattern)
            if elements:
                print(f"  - {pattern}: {len(elements)}개 발견")
        
        # 6. 메타데이터
        print("\n[6] 페이지 메타데이터:")
        title = soup.find('title')
        if title:
            print(f"  페이지 제목: {title.get_text(strip=True)}")
        
        meta_desc = soup.find('meta', attrs={'name': 'description'})
        if meta_desc:
            print(f"  설명: {meta_desc.get('content', '')[:100]}")
        
        # 7. JavaScript 기반 콘텐츠 감지
        print("\n[7] 동적 콘텐츠 감지:")
        scripts = soup.find_all('script')
        print(f"  스크립트 태그: {len(scripts)}개")
        
        # React, Vue, Angular 등 감지
        frameworks = ['react', 'vue', 'angular', 'next', 'nuxt']
        for script in scripts:
            script_text = script.get_text().lower()
            for framework in frameworks:
                if framework in script_text:
                    print(f"  ⚠ {framework.upper()} 프레임워크 감지됨 (동적 렌더링 가능성)")
                    break
        
        return soup
        
    except Exception as e:
        print(f"❌ 오류 발생: {e}")
        return None


def find_specific_content(soup: BeautifulSoup, content_type: str):
    """특정 콘텐츠 타입의 상세 분석"""
    print(f"\n{'='*70}")
    print(f"🔍 '{content_type}' 콘텐츠 상세 분석")
    print(f"{'='*70}")
    
    # 모든 텍스트에서 키워드 검색
    all_text = soup.get_text()
    
    keywords = {
        'report': ['보고서', '연구', 'report', 'research', '발간'],
        'project': ['프로젝트', '과제', 'project', '연구과제'],
        'event': ['행사', '이벤트', 'event', '세미나', '워크샵']
    }
    
    if content_type in keywords:
        for keyword in keywords[content_type]:
            count = all_text.lower().count(keyword.lower())
            if count > 0:
                print(f"  '{keyword}' 키워드: {count}회 발견")


def main():
    """메인 분석 실행"""
    print("="*70)
    print("🌐 GESI 홈페이지 구조 분석 도구")
    print("="*70)
    
    base_url = "https://gesi.kr"
    
    # 1. 메인 페이지 분석
    main_soup = analyze_page_structure(base_url, "메인 페이지")
    time.sleep(1)
    
    # 2. 주요 섹션 URL 추정 및 분석
    possible_sections = [
        ('library', '/library'),
        ('연구', '/research'),
        ('발간물', '/publication'),
        ('보고서', '/report'),
        ('프로젝트', '/project'),
        ('과제', '/projects'),
        ('행사', '/event'),
        ('활동', '/activity'),
        ('소식', '/news'),
    ]
    
    valid_sections = []
    
    print("\n" + "="*70)
    print("📑 주요 섹션 탐색")
    print("="*70)
    
    for name, path in possible_sections:
        url = base_url + path
        print(f"\n시도 중: {url}")
        
        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            response = requests.head(url, headers=headers, timeout=5, allow_redirects=True)
            
            if response.status_code == 200:
                print(f"  ✓ 유효한 페이지 발견!")
                valid_sections.append((name, url))
                time.sleep(1)
                analyze_page_structure(url, f"{name} 페이지")
            else:
                print(f"  ✗ 페이지 없음 (상태 코드: {response.status_code})")
        except Exception as e:
            print(f"  ✗ 접근 실패: {e}")
        
        time.sleep(0.5)
    
    # 3. 결과 요약
    print("\n" + "="*70)
    print("📊 분석 결과 요약")
    print("="*70)
    
    if valid_sections:
        print("\n✓ 발견된 유효 섹션:")
        for name, url in valid_sections:
            print(f"  • {name}: {url}")
    else:
        print("\n⚠ 자동으로 발견된 섹션이 없습니다.")
        print("  홈페이지를 직접 확인하여 URL 구조를 파악해주세요.")
    
    print("\n" + "="*70)
    print("💡 다음 단계:")
    print("  1. 위 분석 결과를 바탕으로 스크래퍼 코드를 맞춤 작성")
    print("  2. 실제 홈페이지를 브라우저에서 확인하여 구조 검증")
    print("  3. 필요시 Selenium을 사용한 동적 콘텐츠 스크래핑 고려")
    print("="*70)
    
    # 결과를 JSON 파일로 저장
    result = {
        'base_url': base_url,
        'valid_sections': valid_sections,
        'analysis_date': time.strftime('%Y-%m-%d %H:%M:%S')
    }
    
    with open('gesi_structure_analysis.json', 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    print("\n✓ 분석 결과가 'gesi_structure_analysis.json'에 저장되었습니다.")


if __name__ == "__main__":
    main()

