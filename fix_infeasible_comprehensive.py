import pandas as pd
import numpy as np
import os
from datetime import datetime

def fix_infeasible_comprehensive():
    """Infeasible 문제를 종합적으로 해결하는 함수"""
    
    print("=== Infeasible 문제 종합 해결 ===")
    
    # 백업 파일 생성
    backup_filename = f"integrated_input_data_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    try:
        # 원본 파일을 백업
        import shutil
        shutil.copy2('integrated_input_data.xlsx', backup_filename)
        print(f"백업 파일 생성: {backup_filename}")
        
        # 데이터 로드
        input_data = {}
        with pd.ExcelFile('integrated_input_data.xlsx') as xls:
            for sheet in ['buses', 'generators', 'loads', 'stores', 'links', 'lines', 'constraints']:
                if sheet in xls.sheet_names:
                    input_data[sheet] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet)
                    print(f"{sheet} 시트 로드: {len(input_data[sheet])}개 행")
        
        # 1. 슬랙 발전기 추가
        print("\n1. 슬랙 발전기 추가 중...")
        generators_df = input_data['generators'].copy()
        buses_df = input_data['buses'].copy()
        
        # 전력 버스별로 슬랙 발전기 추가
        electric_buses = buses_df[buses_df['name'].str.endswith('_EL')]['name'].tolist()
        slack_generators = []
        
        for bus in electric_buses:
            # 해당 버스의 현재 발전기 확인
            current_gens = generators_df[generators_df['bus'] == bus]
            
            # 슬랙 발전기가 없으면 추가
            slack_name = f"{bus.replace('_EL', '')}_Slack_Gen"
            if slack_name not in generators_df['name'].values:
                slack_gen = {
                    'name': slack_name,
                    'bus': bus,
                    'carrier': 'electricity',
                    'p_nom': 0.0,
                    'p_nom_extendable': True,
                    'p_nom_min': 0.0,
                    'p_nom_max': 50000.0,  # 충분히 큰 값
                    'capital_cost': 0.0,
                    'marginal_cost': 1e9,  # 매우 높은 비용
                    'efficiency': 1.0
                }
                slack_generators.append(slack_gen)
        
        if slack_generators:
            slack_df = pd.DataFrame(slack_generators)
            generators_df = pd.concat([generators_df, slack_df], ignore_index=True)
            print(f"슬랙 발전기 {len(slack_generators)}개 추가됨")
        
        # 2. 지역별 백업 발전기 추가 (부족 지역)
        print("\n2. 부족 지역 백업 발전기 추가 중...")
        loads_df = input_data['loads']
        
        # 지역별 수급 분석
        regions = set(loads_df['name'].str.split('_').str[0])
        backup_generators = []
        
        for region in regions:
            region_loads = loads_df[loads_df['name'].str.startswith(f"{region}_")]
            region_gens = generators_df[generators_df['name'].str.startswith(f"{region}_")]
            
            if not region_loads.empty:
                total_load = region_loads['p_set'].sum()
                total_gen = region_gens['p_nom'].sum() if not region_gens.empty else 0
                
                if total_load > total_gen:
                    shortage = total_load - total_gen
                    backup_name = f"{region}_Backup_LNG"
                    backup_bus = f"{region}_EL"
                    
                    # 백업 발전기가 없으면 추가
                    if backup_name not in generators_df['name'].values:
                        backup_gen = {
                            'name': backup_name,
                            'bus': backup_bus,
                            'carrier': 'gas',
                            'p_nom': shortage * 1.5,  # 50% 여유
                            'p_nom_extendable': True,
                            'p_nom_min': 0.0,
                            'p_nom_max': shortage * 3.0,
                            'capital_cost': 500000.0,
                            'marginal_cost': 80.0,  # 일반적인 LNG 비용
                            'efficiency': 0.45
                        }
                        backup_generators.append(backup_gen)
                        print(f"{region} 지역 백업 발전기 추가: {shortage:.0f} MW 부족 -> {shortage*1.5:.0f} MW 설치")
        
        if backup_generators:
            backup_df = pd.DataFrame(backup_generators)
            generators_df = pd.concat([generators_df, backup_df], ignore_index=True)
            print(f"백업 발전기 {len(backup_generators)}개 추가됨")
        
        # 3. CHP 링크 효율 검증 및 수정
        print("\n3. CHP 링크 효율 검증 중...")
        links_df = input_data['links'].copy()
        
        chp_links = links_df[links_df['name'].str.contains('CHP', na=False)]
        for idx, link in chp_links.iterrows():
            # efficiency1 (전력 효율) 확인
            if pd.isna(link.get('efficiency1')) or link.get('efficiency1') == 0:
                links_df.at[idx, 'efficiency1'] = 0.35  # 일반적인 CHP 전력 효율
                print(f"{link['name']} 전력 효율 설정: 0.35")
            
            # efficiency2 (열 효율) 확인  
            if pd.isna(link.get('efficiency2')) or link.get('efficiency2') == 0:
                links_df.at[idx, 'efficiency2'] = 0.40  # 일반적인 CHP 열 효율
                print(f"{link['name']} 열 효율 설정: 0.40")
        
        # 4. 발전기 확장성 설정
        print("\n4. 발전기 확장성 설정 중...")
        
        # 재생에너지와 LNG 발전기의 확장성 활성화
        renewable_types = ['wind', 'solar', 'pv']
        thermal_types = ['gas', 'lng']
        
        for idx, gen in generators_df.iterrows():
            gen_name = str(gen['name']).lower()
            carrier = str(gen.get('carrier', '')).lower()
            
            # 재생에너지 발전기
            if any(rtype in gen_name or rtype in carrier for rtype in renewable_types):
                generators_df.at[idx, 'p_nom_extendable'] = True
                if pd.isna(gen.get('p_nom_max')):
                    generators_df.at[idx, 'p_nom_max'] = gen['p_nom'] * 10  # 10배까지 확장
            
            # 가스/LNG 발전기
            elif any(ttype in gen_name or ttype in carrier for ttype in thermal_types):
                generators_df.at[idx, 'p_nom_extendable'] = True
                if pd.isna(gen.get('p_nom_max')):
                    generators_df.at[idx, 'p_nom_max'] = gen['p_nom'] * 3  # 3배까지 확장
        
        # 5. CO2 제약 비활성화 확인
        print("\n5. CO2 제약 확인 중...")
        constraints_df = input_data['constraints'].copy()
        
        co2_constraints = constraints_df[constraints_df['name'].str.contains('CO2', na=False)]
        if not co2_constraints.empty:
            for idx, constraint in co2_constraints.iterrows():
                current_limit = constraint.get('constant', 0)
                if current_limit < 1e9:  # 10억 톤보다 작으면
                    constraints_df.at[idx, 'constant'] = 1e12  # 1조 톤으로 설정 (실질적 무제한)
                    print(f"CO2 제약 완화: {current_limit} -> 1e12")
        
        # 6. 저장장치 확장성 설정
        print("\n6. 저장장치 확장성 설정 중...")
        if 'stores' in input_data:
            stores_df = input_data['stores'].copy()
            
            for idx, store in stores_df.iterrows():
                stores_df.at[idx, 'e_nom_extendable'] = True
                if pd.isna(store.get('e_nom_max')):
                    stores_df.at[idx, 'e_nom_max'] = store['e_nom'] * 5  # 5배까지 확장
        
        # 7. 수정된 데이터 저장
        print("\n7. 수정된 데이터 저장 중...")
        
        with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
            # 기본 데이터 저장
            generators_df.to_excel(writer, sheet_name='generators', index=False)
            links_df.to_excel(writer, sheet_name='links', index=False)
            constraints_df.to_excel(writer, sheet_name='constraints', index=False)
            
            if 'stores' in input_data:
                stores_df.to_excel(writer, sheet_name='stores', index=False)
            
            # 나머지 시트들은 그대로 복사
            for sheet_name in ['buses', 'loads', 'lines', 'timeseries', 'renewable_patterns', 'load_patterns']:
                if sheet_name in input_data:
                    input_data[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
        
        print("\n=== 수정 완료 ===")
        print(f"총 {len(slack_generators)} 개의 슬랙 발전기 추가")
        print(f"총 {len(backup_generators)} 개의 백업 발전기 추가") 
        print("CHP 링크 효율 설정 완료")
        print("발전기 확장성 설정 완료")
        print("CO2 제약 완화 완료")
        print("저장장치 확장성 설정 완료")
        print("\nInfeasible 문제가 해결되었을 것입니다.")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    fix_infeasible_comprehensive() 