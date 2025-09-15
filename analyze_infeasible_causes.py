import pandas as pd
import numpy as np

def analyze_infeasible_causes():
    print("=== Infeasible 원인 심화 분석 ===")
    
    try:
        # 데이터 로드
        generators_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
        loads_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='loads')
        links_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='links')
        stores_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='stores')
        lines_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='lines')
        
        print("1. === 전체 수급 균형 분석 ===")
        
        # 총 수요 vs 총 공급
        total_demand = loads_df['p_set'].sum()
        total_generation = generators_df['p_nom'].sum()
        slack_generators = generators_df[generators_df['name'].str.contains('Slack|Power_Slack', na=False)]
        total_generation_no_slack = generators_df[~generators_df['name'].str.contains('Slack|Power_Slack', na=False)]['p_nom'].sum()
        
        print(f"총 전력 수요: {total_demand:,.1f} MW")
        print(f"총 발전 용량 (슬랙 포함): {total_generation:,.1f} MW")
        print(f"총 발전 용량 (슬랙 제외): {total_generation_no_slack:,.1f} MW")
        print(f"슬랙 발전기: {len(slack_generators)}개")
        
        balance = total_generation_no_slack - total_demand
        print(f"수급 균형 (슬랙 제외): {balance:,.1f} MW ({'여유' if balance >= 0 else '부족'})")
        
        print("\n2. === 링크 용량 분석 ===")
        
        # 용량이 0인 링크들 확인
        zero_capacity_links = links_df[links_df['p_nom'] == 0]
        print(f"용량이 0인 링크: {len(zero_capacity_links)}개")
        if len(zero_capacity_links) > 0:
            print("용량 0인 링크 유형별 분포:")
            for link_type in ['HP', 'CHP', 'Electrolyser']:
                count = len(zero_capacity_links[zero_capacity_links['name'].str.contains(link_type, na=False)])
                if count > 0:
                    print(f"  - {link_type}: {count}개")
        
        # CHP 전력 생산 잠재력
        chp_links = links_df[links_df['name'].str.contains('CHP', na=False)]
        total_chp_capacity = chp_links['p_nom'].sum()
        print(f"\nCHP 총 용량: {total_chp_capacity:,.1f} MW")
        
        print("\n3. === 지역별 CHP 의존도 분석 ===")
        
        power_buses = ['BSN_EL', 'CBD_EL', 'CND_EL', 'DGU_EL', 'DJN_EL', 'GBD_EL', 'GGD_EL', 
                      'GND_EL', 'GWD_EL', 'GWJ_EL', 'ICN_EL', 'JBD_EL', 'JJD_EL', 'JND_EL', 
                      'SEL_EL', 'SJN_EL', 'USN_EL']
        
        for bus in power_buses:
            region = bus.replace('_EL', '')
            
            # 지역 수요
            region_demand = loads_df[loads_df['bus'] == bus]['p_set'].sum()
            
            # 지역 일반 발전기 (슬랙 제외)
            region_gens = generators_df[
                (generators_df['bus'] == bus) & 
                (~generators_df['name'].str.contains('Slack|Power_Slack', na=False))
            ]
            region_gen_capacity = region_gens['p_nom'].sum()
            
            # 지역 CHP 용량
            region_chp = chp_links[chp_links['name'].str.startswith(region)]
            region_chp_capacity = region_chp['p_nom'].sum() if not region_chp.empty else 0
            
            total_local_supply = region_gen_capacity + region_chp_capacity
            dependency = (region_demand - total_local_supply) / region_demand * 100 if region_demand > 0 else 0
            
            if dependency > 50:  # 50% 이상 외부 의존
                print(f"{region}: 수요 {region_demand:.1f}MW, 일반발전 {region_gen_capacity:.1f}MW, CHP {region_chp_capacity:.1f}MW, 외부의존 {dependency:.1f}%")
        
        print("\n4. === 시간별 패턴 제약 분석 ===")
        
        # 부하 패턴 확인
        try:
            load_patterns = pd.read_excel('integrated_input_data.xlsx', sheet_name='load_patterns')
            print(f"부하 패턴 데이터: {load_patterns.shape}")
            
            # 패턴의 최대값과 최소값 비교
            if len(load_patterns.columns) > 3:
                el_pattern = pd.to_numeric(load_patterns.iloc[7:, 1], errors='coerce').dropna()
                if len(el_pattern) > 0:
                    pattern_max = el_pattern.max()
                    pattern_min = el_pattern.min()
                    pattern_ratio = pattern_max / pattern_min if pattern_min > 0 else float('inf')
                    print(f"전력 패턴 변동: 최대/최소 = {pattern_ratio:.2f}")
                    
                    if pattern_ratio > 3:
                        print("⚠️ 높은 시간별 변동성이 문제 원인일 수 있음")
        except:
            print("부하 패턴 데이터 없음")
        
        print("\n5. === 잠재적 문제 진단 ===")
        
        issues = []
        
        # Issue 1: HP 용량이 모두 0
        hp_links = links_df[links_df['name'].str.contains('HP', na=False)]
        if (hp_links['p_nom'] == 0).all():
            issues.append("❌ 모든 HP (Heat Pump) 용량이 0 → 열 수요 충족 불가능할 수 있음")
        
        # Issue 2: Electrolyser 용량이 모두 0  
        electro_links = links_df[links_df['name'].str.contains('Electrolyser', na=False)]
        if (electro_links['p_nom'] == 0).all():
            issues.append("❌ 모든 Electrolyser 용량이 0 → 수소 수요 충족 불가능할 수 있음")
        
        # Issue 3: 일부 CHP 용량이 0
        zero_chp = chp_links[chp_links['p_nom'] == 0]
        if len(zero_chp) > 0:
            issues.append(f"❌ {len(zero_chp)}개 지역의 CHP 용량이 0 → 해당 지역 열/전력 동시 공급 불가")
        
        # Issue 4: 전체 수급 부족
        if balance < 0:
            issues.append(f"❌ 전체 발전량 부족: {abs(balance):,.1f} MW")
        
        if issues:
            print("발견된 문제들:")
            for issue in issues:
                print(f"  {issue}")
        else:
            print("✅ 명확한 용량 문제는 발견되지 않음")
            print("   → 시간별 제약, 링크 효율, 또는 복합적 제약이 원인일 가능성")
        
        print("\n6. === 권장 해결책 ===")
        print("1. HP와 Electrolyser 용량을 0이 아닌 값으로 설정")
        print("2. 용량이 0인 CHP 지역에 최소 용량 부여")
        print("3. 송전선로 용량을 일시적으로 증가시켜 테스트")
        print("4. 시간별 부하 패턴의 평활화 또는 저장장치 용량 증대")
        
    except Exception as e:
        print(f"분석 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    analyze_infeasible_causes() 