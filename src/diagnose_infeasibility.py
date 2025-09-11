import pandas as pd
import numpy as np
from datetime import datetime

def diagnose_network_feasibility():
    """네트워크 실행 가능성 진단"""
    
    print("=== 네트워크 실행 가능성 진단 ===")
    
    # 데이터 로드
    try:
        xls = pd.ExcelFile('integrated_input_data.xlsx')
        generators_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
        loads_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='loads')
        lines_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='lines')
        
        print(f"발전기 수: {len(generators_df)}")
        print(f"부하 수: {len(loads_df)}")
        print(f"송전선로 수: {len(lines_df)}")
        
    except Exception as e:
        print(f"데이터 로드 오류: {e}")
        return
    
    # 1. 전체 발전 용량 vs 전체 부하 수요 비교
    print("\n=== 1. 전체 발전 용량 vs 부하 수요 ===")
    
    total_generation_capacity = generators_df['p_nom'].sum()
    total_load_demand = loads_df['p_set'].sum()
    
    print(f"총 발전 용량: {total_generation_capacity:,.0f} MW")
    print(f"총 부하 수요: {total_load_demand:,.0f} MW")
    print(f"여유율: {((total_generation_capacity - total_load_demand) / total_load_demand * 100):,.1f}%")
    
    if total_generation_capacity < total_load_demand:
        print("⚠️ 경고: 총 발전 용량이 총 부하 수요보다 부족합니다!")
    
    # 2. 지역별 발전 용량 vs 부하 수요 분석
    print("\n=== 2. 지역별 발전 용량 vs 부하 수요 ===")
    
    # 지역별 발전 용량 계산
    regional_generation = {}
    for _, gen in generators_df.iterrows():
        region = gen['name'].split('_')[0]
        if region not in regional_generation:
            regional_generation[region] = 0
        regional_generation[region] += gen['p_nom']
    
    # 지역별 부하 수요 계산 (전력 부하만)
    regional_demand = {}
    for _, load in loads_df.iterrows():
        if '_Demand_EL' in load['name']:  # 전력 부하만
            region = load['name'].split('_')[0]
            if region not in regional_demand:
                regional_demand[region] = 0
            regional_demand[region] += load['p_set']
    
    # 지역별 비교
    deficit_regions = []
    surplus_regions = []
    
    print(f"{'지역':<6} {'발전용량(MW)':<12} {'부하수요(MW)':<12} {'차이(MW)':<10} {'여유율(%)':<10} {'상태'}")
    print("-" * 70)
    
    for region in sorted(set(list(regional_generation.keys()) + list(regional_demand.keys()))):
        gen_cap = regional_generation.get(region, 0)
        demand = regional_demand.get(region, 0)
        difference = gen_cap - demand
        
        if demand > 0:
            margin = (difference / demand) * 100
        else:
            margin = float('inf') if gen_cap > 0 else 0
        
        status = "부족" if difference < 0 else "충분"
        if difference < 0:
            deficit_regions.append(region)
        else:
            surplus_regions.append(region)
        
        print(f"{region:<6} {gen_cap:<12,.0f} {demand:<12,.0f} {difference:<10,.0f} {margin:<10.1f} {status}")
    
    print(f"\n부족 지역: {deficit_regions}")
    print(f"잉여 지역: {surplus_regions}")
    
    # 3. 송전선로 연결성 분석
    print("\n=== 3. 송전선로 연결성 분석 ===")
    
    # 지역 간 연결 매트릭스 생성
    regions = sorted(set(list(regional_generation.keys()) + list(regional_demand.keys())))
    connection_matrix = pd.DataFrame(0, index=regions, columns=regions)
    
    total_transmission_capacity = 0
    for _, line in lines_df.iterrows():
        bus0_region = line['bus0'].split('_')[0]
        bus1_region = line['bus1'].split('_')[0]
        capacity = line['s_nom']
        total_transmission_capacity += capacity
        
        if bus0_region in regions and bus1_region in regions:
            connection_matrix.loc[bus0_region, bus1_region] += capacity
            connection_matrix.loc[bus1_region, bus0_region] += capacity
    
    print(f"총 송전 용량: {total_transmission_capacity:,.0f} MVA")
    
    # 부족 지역의 송전 연결 확인
    print("\n부족 지역의 송전 연결 상태:")
    for region in deficit_regions:
        if region in connection_matrix.index:
            connections = connection_matrix.loc[region]
            connected_regions = connections[connections > 0]
            total_import_capacity = connected_regions.sum()
            
            deficit_amount = abs(regional_generation.get(region, 0) - regional_demand.get(region, 0))
            
            print(f"\n{region} 지역:")
            print(f"  부족량: {deficit_amount:,.0f} MW")
            print(f"  총 수입 가능 용량: {total_import_capacity:,.0f} MVA")
            print(f"  연결된 지역: {list(connected_regions.index)}")
            
            if total_import_capacity < deficit_amount:
                print(f"  ⚠️ 송전 용량 부족: {deficit_amount - total_import_capacity:,.0f} MW 추가 필요")
    
    # 4. 발전원별 분석
    print("\n=== 4. 발전원별 분석 ===")
    
    generation_by_type = {}
    for _, gen in generators_df.iterrows():
        gen_type = 'Unknown'
        if 'Nuclear' in gen['name']:
            gen_type = '원자력'
        elif 'PV' in gen['name']:
            gen_type = '태양광'
        elif 'WT' in gen['name']:
            gen_type = '풍력'
        elif 'Coal' in gen['name']:
            gen_type = '석탄'
        elif 'LNG' in gen['name']:
            gen_type = 'LNG'
        elif 'Hydro' in gen['name']:
            gen_type = '수력'
        elif 'Other' in gen['name']:
            gen_type = '기타'
        
        if gen_type not in generation_by_type:
            generation_by_type[gen_type] = 0
        generation_by_type[gen_type] += gen['p_nom']
    
    print(f"{'발전원':<8} {'용량(MW)':<12} {'비율(%)':<8}")
    print("-" * 30)
    for gen_type, capacity in sorted(generation_by_type.items(), key=lambda x: x[1], reverse=True):
        ratio = (capacity / total_generation_capacity) * 100
        print(f"{gen_type:<8} {capacity:<12,.0f} {ratio:<8.1f}")
    
    # 5. 해결 방안 제시
    print("\n=== 5. 해결 방안 ===")
    
    if deficit_regions:
        print("부족 지역 해결 방안:")
        for region in deficit_regions:
            deficit_amount = abs(regional_generation.get(region, 0) - regional_demand.get(region, 0))
            print(f"\n{region} 지역 ({deficit_amount:,.0f} MW 부족):")
            print(f"  1. 발전 용량 증설")
            print(f"  2. 송전선로 용량 증대")
            print(f"  3. 부하 감축")
    
    # 6. 최소 필요 송전 용량 계산
    total_deficit = sum(abs(regional_generation.get(region, 0) - regional_demand.get(region, 0)) 
                       for region in deficit_regions)
    
    if total_deficit > 0:
        print(f"\n최소 필요 송전 용량: {total_deficit:,.0f} MW")
        print(f"현재 송전 용량: {total_transmission_capacity:,.0f} MVA")
        
        if total_transmission_capacity < total_deficit:
            additional_needed = total_deficit - total_transmission_capacity
            print(f"추가 필요 송전 용량: {additional_needed:,.0f} MVA")

if __name__ == "__main__":
    diagnose_network_feasibility() 