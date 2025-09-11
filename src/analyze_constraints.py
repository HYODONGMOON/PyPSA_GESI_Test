import pandas as pd
import numpy as np

def analyze_constraints():
    """제약조건 분석"""
    print("=== 제약조건 분석 시작 ===")
    
    # 데이터 로드
    input_data = {}
    xls = pd.ExcelFile('integrated_input_data.xlsx')
    for sheet_name in xls.sheet_names:
        input_data[sheet_name] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet_name)
    
    # 1. 발전량 vs 부하량 분석
    print("\n1. 발전량 vs 부하량 분석")
    
    # 총 발전 용량
    total_generation = input_data['generators']['p_nom'].sum()
    print(f"총 발전 용량: {total_generation:,.0f} MW")
    
    # 총 부하
    total_load = input_data['loads']['p_set'].sum()
    print(f"총 부하: {total_load:,.0f} MW")
    
    # 여유율
    reserve_margin = (total_generation - total_load) / total_load * 100
    print(f"예비율: {reserve_margin:.1f}%")
    
    # 2. 지역별 발전량 vs 부하량
    print("\n2. 지역별 발전량 vs 부하량")
    
    regional_analysis = {}
    
    # 지역 코드 추출
    regions = set()
    for gen_name in input_data['generators']['name']:
        if '_' in str(gen_name):
            region = str(gen_name).split('_')[0]
            regions.add(region)
    
    for region in sorted(regions):
        # 지역별 발전량
        region_gens = input_data['generators'][input_data['generators']['name'].str.startswith(f"{region}_")]
        region_generation = region_gens['p_nom'].sum()
        
        # 지역별 부하 (전력만)
        region_loads = input_data['loads'][input_data['loads']['name'].str.startswith(f"{region}_Demand_EL")]
        region_load = region_loads['p_set'].sum() if not region_loads.empty else 0
        
        # 수급 균형
        balance = region_generation - region_load
        
        regional_analysis[region] = {
            'generation': region_generation,
            'load': region_load,
            'balance': balance
        }
        
        print(f"{region}: 발전 {region_generation:,.0f}MW, 부하 {region_load:,.0f}MW, 균형 {balance:,.0f}MW")
    
    # 3. 부족 지역 식별
    print("\n3. 전력 부족 지역")
    deficit_regions = []
    for region, data in regional_analysis.items():
        if data['balance'] < 0:
            deficit_regions.append(region)
            print(f"{region}: {data['balance']:,.0f}MW 부족")
    
    # 4. 재생에너지 패턴 분석
    print("\n4. 재생에너지 패턴 분석")
    if 'renewable_patterns' in input_data:
        patterns = input_data['renewable_patterns']
        if 'PV' in patterns.columns:
            pv_max = patterns['PV'].max()
            pv_min = patterns['PV'].min()
            pv_avg = patterns['PV'].mean()
            print(f"PV 패턴: 최대 {pv_max:.3f}, 최소 {pv_min:.3f}, 평균 {pv_avg:.3f}")
        
        if 'WT' in patterns.columns:
            wt_max = patterns['WT'].max()
            wt_min = patterns['WT'].min()
            wt_avg = patterns['WT'].mean()
            print(f"WT 패턴: 최대 {wt_max:.3f}, 최소 {wt_min:.3f}, 평균 {wt_avg:.3f}")
    
    # 5. 저장장치 용량 분석
    print("\n5. 저장장치 용량 분석")
    if 'stores' in input_data:
        total_storage = input_data['stores']['e_nom'].sum()
        print(f"총 저장 용량: {total_storage:,.0f} MWh")
        
        # 지역별 저장 용량
        for region in sorted(regions):
            region_stores = input_data['stores'][input_data['stores']['name'].str.startswith(f"{region}_")]
            region_storage = region_stores['e_nom'].sum()
            print(f"{region} 저장 용량: {region_storage:,.0f} MWh")
    
    # 6. 선로 용량 분석
    print("\n6. 선로 용량 분석")
    if 'lines' in input_data:
        total_line_capacity = input_data['lines']['s_nom'].sum()
        print(f"총 선로 용량: {total_line_capacity:,.0f} MVA")
        
        # 부족 지역 연결 선로
        print("\n부족 지역 연결 선로:")
        for region in deficit_regions:
            connected_lines = input_data['lines'][
                (input_data['lines']['bus0'].str.contains(f"{region}_")) |
                (input_data['lines']['bus1'].str.contains(f"{region}_"))
            ]
            total_capacity = connected_lines['s_nom'].sum()
            print(f"{region} 연결 선로 용량: {total_capacity:,.0f} MVA")
    
    # 7. 제약조건 확인
    print("\n7. 제약조건 확인")
    if 'constraints' in input_data:
        constraints = input_data['constraints']
        print("설정된 제약조건:")
        for _, constraint in constraints.iterrows():
            print(f"- {constraint['name']}: {constraint['type']} {constraint['sense']} {constraint['constant']}")
    
    return regional_analysis, deficit_regions

if __name__ == "__main__":
    analyze_constraints() 