"""
Bus Prices (한계가격) 분석 스크립트
"""
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import seaborn as sns
import numpy as np

# 한글 폰트 설정
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

def analyze_bus_prices(result_file):
    """
    Bus Prices 분석
    
    Args:
        result_file: 결과 엑셀 파일 경로
    """
    print("=" * 80)
    print("Bus Prices (한계가격) 분석")
    print("=" * 80)
    
    # 데이터 로드
    try:
        bus_prices = pd.read_excel(result_file, sheet_name='Bus_Prices', index_col=0)
        print(f"\n[OK] Bus_Prices 시트 로드 성공")
        print(f"   - 시간: {len(bus_prices)}시간")
        print(f"   - 버스: {len(bus_prices.columns)}개")
    except Exception as e:
        print(f"\n[실패] Bus_Prices 시트 로드 실패: {str(e)}")
        return
    
    # 지역 대표 전력 버스만 필터링 (bus_로 시작하지 않고 _EL로 끝나는 버스)
    # bus_XXX_YYY 형태는 물리적 송전망 모델링용이므로 제외
    region_codes = ['BSN', 'SEL', 'ICN', 'DGU', 'DJN', 'USN', 'GWJ', 'GWD', 
                    'GBD', 'GGD', 'GND', 'JBD', 'JJD', 'JND', 'CBD', 'CND', 'SJN']
    el_buses = [f"{code}_EL" for code in region_codes if f"{code}_EL" in bus_prices.columns]
    
    bus_prices_el = bus_prices[el_buses]
    
    print(f"\n지역 대표 전력 버스: {len(el_buses)}개")
    print(f"  {', '.join(el_buses)}")
    print(f"\n[참고] 물리적 송전망 버스(bus_XXX_YYY)는 분석에서 제외됨")
    
    # 1. 기초 통계
    print("\n" + "=" * 80)
    print("1. 기초 통계 (원/MWh)")
    print("=" * 80)
    
    stats = pd.DataFrame({
        '평균': bus_prices_el.mean(),
        '중앙값': bus_prices_el.median(),
        '최소': bus_prices_el.min(),
        '최대': bus_prices_el.max(),
        '표준편차': bus_prices_el.std()
    }).round(0)
    
    print(stats)
    
    # 2. 지역별 평균 가격 순위
    print("\n" + "=" * 80)
    print("2. 지역별 평균 전력가격 순위 (높은 순)")
    print("=" * 80)
    
    avg_prices = bus_prices_el.mean().sort_values(ascending=False)
    for i, (bus, price) in enumerate(avg_prices.items(), 1):
        region = bus.replace('_EL', '')
        print(f"{i:2d}. {region:5s}: {price:>10,.0f} 원/MWh")
    
    # 3. 시간대별 평균 가격
    print("\n" + "=" * 80)
    print("3. 시간대별 평균 전력가격")
    print("=" * 80)
    
    bus_prices_el['시간'] = pd.to_datetime(bus_prices_el.index).hour
    hourly_avg = bus_prices_el.groupby('시간').mean().mean(axis=1)
    
    print("\n시간 | 평균가격 (원/MWh)")
    print("-" * 30)
    for hour, price in hourly_avg.items():
        bar = '█' * int(price / 5000)
        print(f"{hour:2d}시 | {price:>10,.0f} {bar}")
    
    # 4. 가격 변동성이 큰 지역 TOP 5
    print("\n" + "=" * 80)
    print("4. 가격 변동성이 큰 지역 TOP 5")
    print("=" * 80)
    
    volatility = bus_prices_el.drop('시간', axis=1).std().sort_values(ascending=False).head(5)
    for i, (bus, std) in enumerate(volatility.items(), 1):
        region = bus.replace('_EL', '')
        cv = std / bus_prices_el[bus].mean() * 100  # 변동계수
        print(f"{i}. {region:5s}: 표준편차 {std:>10,.0f} 원/MWh (CV: {cv:.1f}%)")
    
    # 5. 지역 간 가격 차이 (송전망 혼잡도)
    print("\n" + "=" * 80)
    print("5. 지역 간 최대 가격 차이 (송전망 혼잡도)")
    print("=" * 80)
    
    bus_prices_clean = bus_prices_el.drop('시간', axis=1)
    max_diff = bus_prices_clean.max(axis=1) - bus_prices_clean.min(axis=1)
    
    print(f"평균 최대 가격차: {max_diff.mean():>10,.0f} 원/MWh")
    print(f"최대 가격차:     {max_diff.max():>10,.0f} 원/MWh")
    print(f"최소 가격차:     {max_diff.min():>10,.0f} 원/MWh")
    
    # 최대 가격차가 발생한 시간
    max_diff_idx = max_diff.idxmax()
    max_diff_time = bus_prices_clean.loc[max_diff_idx]
    max_region = max_diff_time.idxmax().replace('_EL', '')
    min_region = max_diff_time.idxmin().replace('_EL', '')
    
    print(f"\n최대 가격차 발생 시간: {max_diff_idx}")
    print(f"  - 최고가 지역: {max_region} ({max_diff_time.max():,.0f} 원/MWh)")
    print(f"  - 최저가 지역: {min_region} ({max_diff_time.min():>,.0f} 원/MWh)")
    print(f"  - 가격차: {max_diff.max():,.0f} 원/MWh")
    
    # 6. 그래프 생성
    print("\n" + "=" * 80)
    print("6. 시각화 생성 중...")
    print("=" * 80)
    
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    
    # 6-1. 시간별 주요 지역 가격
    ax1 = axes[0, 0]
    major_regions = ['SEL_EL', 'BSN_EL', 'ICN_EL', 'DGU_EL', 'GWJ_EL']
    for region in major_regions:
        if region in bus_prices_clean.columns:
            ax1.plot(bus_prices_clean.index, bus_prices_clean[region]/1000, 
                    label=region.replace('_EL', ''), alpha=0.7)
    ax1.set_xlabel('시간')
    ax1.set_ylabel('전력가격 (천원/MWh)')
    ax1.set_title('주요 지역별 전력가격 추이')
    ax1.legend()
    ax1.grid(True, alpha=0.3)
    
    # 6-2. 지역별 평균 가격 막대 그래프
    ax2 = axes[0, 1]
    avg_prices_sorted = avg_prices.sort_values()
    colors = plt.cm.RdYlGn_r(np.linspace(0.2, 0.8, len(avg_prices_sorted)))
    avg_prices_sorted.plot(kind='barh', ax=ax2, color=colors)
    ax2.set_xlabel('평균 전력가격 (원/MWh)')
    ax2.set_title('지역별 평균 전력가격')
    ax2.yaxis.set_label_text('')
    
    # 6-3. 시간대별 평균 가격
    ax3 = axes[1, 0]
    hourly_avg.plot(kind='bar', ax=ax3, color='steelblue')
    ax3.set_xlabel('시간')
    ax3.set_ylabel('평균 전력가격 (원/MWh)')
    ax3.set_title('시간대별 평균 전력가격')
    ax3.grid(True, alpha=0.3, axis='y')
    
    # 6-4. 가격 분포 박스플롯
    ax4 = axes[1, 1]
    bus_prices_clean.boxplot(ax=ax4, rot=90)
    ax4.set_ylabel('전력가격 (원/MWh)')
    ax4.set_title('지역별 전력가격 분포')
    ax4.grid(True, alpha=0.3, axis='y')
    
    plt.tight_layout()
    
    # 저장
    output_file = result_file.replace('.xlsx', '_bus_prices_analysis.png')
    plt.savefig(output_file, dpi=150, bbox_inches='tight')
    print(f"\n[OK] 시각화 저장: {output_file}")
    
    plt.show()
    
    print("\n" + "=" * 80)
    print("분석 완료!")
    print("=" * 80)


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        result_file = sys.argv[1]
    else:
        # 최신 결과 파일 자동 검색
        import os
        import glob
        
        result_dirs = glob.glob('results/rolling_horizon_*/rolling_horizon_result_*.xlsx')
        if not result_dirs:
            result_dirs = glob.glob('results/*/optimization_result_*.xlsx')
        
        if result_dirs:
            result_file = max(result_dirs, key=os.path.getmtime)
            print(f"최신 결과 파일 사용: {result_file}")
        else:
            print("결과 파일을 찾을 수 없습니다.")
            print("사용법: python analyze_bus_prices.py <결과파일.xlsx>")
            sys.exit(1)
    
    analyze_bus_prices(result_file)
