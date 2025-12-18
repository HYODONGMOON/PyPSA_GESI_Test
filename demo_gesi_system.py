# -*- coding: utf-8 -*-
"""
GESI Annual Report 시스템 데모 스크립트
간단한 테스트 및 데이터 확인용
"""

from gesi_annual_report_system import GESIAnnualReportCollector, GESIDatabase
from datetime import datetime


def demo_basic_usage():
    """기본 사용법 데모"""
    print("\n" + "="*70)
    print("📋 데모 1: 기본 데이터 수집")
    print("="*70)
    
    # 초기화 (브라우저 보이게 설정)
    collector = GESIAnnualReportCollector(headless=True)
    
    try:
        # Library 1페이지만 수집 (테스트용)
        print("\n🔍 Library 페이지 1개만 수집합니다...")
        result = collector.update_library(max_pages=1)
        print(f"결과: 총 {result['total']}개 중 {result['new']}개 신규 추가")
        
    finally:
        collector.close()


def demo_database_query():
    """데이터베이스 조회 데모"""
    print("\n" + "="*70)
    print("📊 데모 2: 데이터베이스 조회")
    print("="*70)
    
    db = GESIDatabase()
    
    try:
        # 전체 보고서 조회
        all_reports = db.get_library()
        print(f"\n📚 전체 보고서: {len(all_reports)}건")
        
        if not all_reports.empty:
            print("\n최근 보고서 5건:")
            for idx, row in all_reports.head(5).iterrows():
                print(f"  {idx+1}. [{row['category']}] {row['title']}")
                print(f"     저자: {row['author']}, 날짜: {row['published_date']}")
        
        # 연도별 통계
        current_year = datetime.now().year
        for year in [current_year, current_year-1]:
            reports_year = db.get_library(year=year)
            projects_year = db.get_projects(year=year)
            events_year = db.get_events(year=year)
            
            print(f"\n📅 {year}년 통계:")
            print(f"   보고서: {len(reports_year)}건")
            print(f"   프로젝트: {len(projects_year)}건")
            print(f"   행사: {len(events_year)}건")
        
        # 카테고리별 분포
        if not all_reports.empty and 'category' in all_reports.columns:
            print("\n📊 카테고리별 보고서 분포:")
            category_counts = all_reports['category'].value_counts()
            for category, count in category_counts.items():
                print(f"   {category}: {count}건")
        
    finally:
        db.close()


def demo_export_report():
    """Annual Report 생성 데모"""
    print("\n" + "="*70)
    print("📄 데모 3: Annual Report Excel 생성")
    print("="*70)
    
    collector = GESIAnnualReportCollector(headless=True)
    
    try:
        current_year = datetime.now().year
        
        # Excel 파일 생성
        output_path = collector.export_to_excel(
            year=current_year,
            output_dir="demo_output"
        )
        
        print(f"\n✅ Excel 파일이 생성되었습니다:")
        print(f"   {output_path}")
        
    finally:
        collector.close()


def demo_incremental_update():
    """점진적 업데이트 데모"""
    print("\n" + "="*70)
    print("🔄 데모 4: 점진적 업데이트 (중복 방지)")
    print("="*70)
    
    collector = GESIAnnualReportCollector(headless=True)
    
    try:
        print("\n첫 번째 수집:")
        result1 = collector.update_library(max_pages=1)
        print(f"  총 {result1['total']}개 중 {result1['new']}개 신규 추가")
        
        print("\n두 번째 수집 (동일 페이지):")
        result2 = collector.update_library(max_pages=1)
        print(f"  총 {result2['total']}개 중 {result2['new']}개 신규 추가")
        print(f"  → 중복이 자동으로 필터링되었습니다!")
        
    finally:
        collector.close()


def main():
    """데모 메인"""
    print("""
╔═══════════════════════════════════════════════════════════════╗
║                                                               ║
║         GESI Annual Report 시스템 데모                         ║
║                                                               ║
╚═══════════════════════════════════════════════════════════════╝

이 데모는 GESI Annual Report 시스템의 주요 기능을 보여줍니다.
    """)
    
    demos = [
        ("기본 데이터 수집", demo_basic_usage),
        ("데이터베이스 조회", demo_database_query),
        ("Annual Report 생성", demo_export_report),
        ("점진적 업데이트", demo_incremental_update)
    ]
    
    while True:
        print("\n" + "="*70)
        print("실행할 데모를 선택하세요:")
        print("="*70)
        
        for i, (name, _) in enumerate(demos, 1):
            print(f"  {i}. {name}")
        print(f"  0. 종료")
        print()
        
        try:
            choice = input("선택 (0-4): ").strip()
            
            if choice == '0':
                print("\n👋 데모를 종료합니다.")
                break
            
            choice_num = int(choice)
            if 1 <= choice_num <= len(demos):
                demos[choice_num - 1][1]()
            else:
                print("❌ 잘못된 선택입니다.")
                
        except ValueError:
            print("❌ 숫자를 입력해주세요.")
        except KeyboardInterrupt:
            print("\n\n👋 사용자가 중단했습니다.")
            break
        except Exception as e:
            print(f"\n❌ 오류 발생: {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    main()

