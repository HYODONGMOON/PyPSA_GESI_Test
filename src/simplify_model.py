import pandas as pd
import os
import shutil
from datetime import datetime

def simplify_model():
    """
    최적화 문제를 해결하기 위해 모델을 크게 단순화하는 함수
    
    1. 모든 제약조건 제거
    2. 네트워크 구조 단순화
    3. 단일 운영 시점만 고려 (시간 시리즈 무시)
    """
    print("모델 단순화 중...")
    
    # 파일 경로
    input_file = 'integrated_input_data.xlsx'
    simple_model_file = 'simplified_input_data.xlsx'
    
    # 백업 파일 생성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = f'integrated_input_data_backup_simple_{timestamp}.xlsx'
    
    if not os.path.exists(input_file):
        print(f"오류: {input_file} 파일을 찾을 수 없습니다.")
        return False
    
    # 백업 생성
    shutil.copy2(input_file, backup_file)
    print(f"원본 파일을 {backup_file}로 백업했습니다.")
    
    try:
        # 데이터 로드
        with pd.ExcelFile(input_file) as xls:
            buses = pd.read_excel(xls, sheet_name='buses')
            links = pd.read_excel(xls, sheet_name='links', nrows=None)
            generators = pd.read_excel(xls, sheet_name='generators')
            loads = pd.read_excel(xls, sheet_name='loads')
            lines = pd.read_excel(xls, sheet_name='lines')
            stores = pd.read_excel(xls, sheet_name='stores')
        
        print("기존 데이터 정보:")
        print(f"버스: {len(buses)}개")
        print(f"발전기: {len(generators)}개")
        print(f"부하: {len(loads)}개")
        print(f"선로: {len(lines)}개")
        print(f"링크: {len(links)}개")
        print(f"저장장치: {len(stores)}개")
        
        # 1. 가장 단순한 모델로 축소: 외부 그리드 + 각 지역의 대표 버스 + 발전기 + 부하
        # 대표 버스 선택
        simple_buses = []
        region_codes = set()
        
        # 지역 코드 추출
        for bus in buses['name']:
            if '_EL' in bus:
                region_code = bus.split('_')[0]
                region_codes.add(region_code)
        
        print(f"\n발견된 지역 코드: {', '.join(sorted(region_codes))}")
        
        # 각 지역별 전력 버스만 사용
        for region in sorted(region_codes):
            buses_in_region = buses[buses['name'].str.startswith(f"{region}_")]
            el_buses = buses_in_region[buses_in_region['name'].str.contains('_EL')]
            
            if not el_buses.empty:
                simple_buses.append(el_buses.iloc[0].to_dict())
        
        # External_Grid와 Slack_Bus 추가
        external_bus = buses[buses['name'] == 'External_Grid']
        if not external_bus.empty:
            simple_buses.append(external_bus.iloc[0].to_dict())
        else:
            simple_buses.append({
                'name': 'External_Grid',
                'v_nom': 345,
                'carrier': 'AC',
                'x': 0,
                'y': 0
            })
        
        slack_bus = buses[buses['name'] == 'Slack_Bus']
        if not slack_bus.empty:
            simple_buses.append(slack_bus.iloc[0].to_dict())
        else:
            simple_buses.append({
                'name': 'Slack_Bus',
                'v_nom': 345,
                'carrier': 'AC',
                'x': 0,
                'y': 0
            })
        
        # 2. 각 지역별 총 부하 집계
        simple_loads = []
        for region in sorted(region_codes):
            region_buses = [bus['name'] for bus in simple_buses if bus['name'].startswith(f"{region}_")]
            region_loads = loads[loads['bus'].isin(region_buses)]
            
            if not region_loads.empty:
                total_load = region_loads['p_set'].sum()
                
                # 간소화된 부하 생성
                simple_load = {
                    'name': f"{region}_Load",
                    'bus': f"{region}_EL",
                    'p_set': total_load,
                    'carrier': 'AC'
                }
                simple_loads.append(simple_load)
        
        # Slack_Bus에 작은 부하 추가
        simple_loads.append({
            'name': 'Slack_Load',
            'bus': 'Slack_Bus',
            'p_set': 0.1,
            'carrier': 'AC'
        })
        
        # 3. 각 지역별 발전기 집계
        simple_generators = []
        for region in sorted(region_codes):
            region_buses = [bus['name'] for bus in simple_buses if bus['name'].startswith(f"{region}_")]
            region_gens = generators[generators['bus'].isin(region_buses)]
            
            # 화력 발전
            thermal_gens = region_gens[region_gens['carrier'].isin(['coal', 'gas', 'oil', 'nuclear'])]
            if not thermal_gens.empty:
                thermal_capacity = thermal_gens['p_nom'].sum()
                simple_generators.append({
                    'name': f"{region}_Thermal",
                    'bus': f"{region}_EL",
                    'carrier': 'thermal',
                    'p_nom': thermal_capacity * 1.2,  # 20% 여유 추가
                    'p_nom_extendable': True,
                    'p_nom_min': 0,
                    'p_nom_max': thermal_capacity * 5,
                    'marginal_cost': 50
                })
            
            # 재생에너지 발전
            renewables = region_gens[region_gens['carrier'].isin(['wind', 'solar', 'hydro', 'biomass'])]
            if not renewables.empty:
                renewable_capacity = renewables['p_nom'].sum()
                simple_generators.append({
                    'name': f"{region}_Renewable",
                    'bus': f"{region}_EL",
                    'carrier': 'renewable',
                    'p_nom': renewable_capacity * 1.2,  # 20% 여유 추가
                    'p_nom_extendable': True,
                    'p_nom_min': 0,
                    'p_nom_max': renewable_capacity * 10,
                    'marginal_cost': 10
                })
        
        # Slack 발전기 추가
        simple_generators.append({
            'name': 'Slack_Generator',
            'bus': 'Slack_Bus',
            'carrier': 'slack',
            'p_nom': 100000,
            'p_nom_extendable': True,
            'p_nom_min': 0,
            'p_nom_max': 100000,
            'marginal_cost': 10000
        })
        
        # 4. 단순 링크 생성 (모든 지역을 External_Grid에 연결)
        simple_links = []
        
        # 지역간 연결 로직 간소화 (각 지역을 External_Grid에 연결)
        for region in sorted(region_codes):
            # 수입 링크
            simple_links.append({
                'name': f"Import_{region}",
                'bus0': 'External_Grid',
                'bus1': f"{region}_EL",
                'p_nom': 5000,
                'p_nom_extendable': True,
                'p_nom_min': 0,
                'p_nom_max': 10000,
                'efficiency': 1.0,
                'marginal_cost': 100
            })
            
            # 수출 링크
            simple_links.append({
                'name': f"Export_{region}",
                'bus0': f"{region}_EL",
                'bus1': 'External_Grid',
                'p_nom': 5000,
                'p_nom_extendable': True,
                'p_nom_min': 0,
                'p_nom_max': 10000,
                'efficiency': 0.95,
                'marginal_cost': -10
            })
        
        # Slack 버스 연결
        simple_links.append({
            'name': 'External_To_Slack',
            'bus0': 'External_Grid',
            'bus1': 'Slack_Bus',
            'p_nom': 100000,
            'p_nom_extendable': True,
            'p_nom_min': 0,
            'p_nom_max': 100000,
            'efficiency': 1.0,
            'marginal_cost': 0.1
        })
        
        # 5. 단순 선로 생성 (인접 지역만 연결)
        simple_lines = []
        
        # 기존 지역간 연결 정보 사용
        existing_connections = set()
        for _, line in lines.iterrows():
            if pd.notna(line['bus0']) and pd.notna(line['bus1']):
                bus0 = line['bus0']
                bus1 = line['bus1']
                
                # 지역 코드 추출
                if '_' in bus0 and '_' in bus1:
                    region0 = bus0.split('_')[0]
                    region1 = bus1.split('_')[0]
                    
                    if region0 in region_codes and region1 in region_codes and region0 != region1:
                        connection = tuple(sorted([region0, region1]))
                        existing_connections.add(connection)
        
        # 단순 라인 생성
        for idx, (region1, region2) in enumerate(existing_connections):
            line_data = {
                'name': f"Line_{region1}_{region2}",
                'bus0': f"{region1}_EL",
                'bus1': f"{region2}_EL",
                'carrier': 'AC',
                'x': 0.2,
                'r': 0.05,
                's_nom': 2000,
                's_nom_extendable': True,
                's_nom_min': 0,
                's_nom_max': 10000,
                'length': 100,
                'v_nom': 345
            }
            simple_lines.append(line_data)
        
        print(f"\n단순화된 모델 정보:")
        print(f"버스: {len(simple_buses)}개")
        print(f"발전기: {len(simple_generators)}개")
        print(f"부하: {len(simple_loads)}개")
        print(f"선로: {len(simple_lines)}개")
        print(f"링크: {len(simple_links)}개")
        
        # 간소화된 모델을 별도 파일로 저장
        with pd.ExcelWriter(simple_model_file) as writer:
            pd.DataFrame(simple_buses).to_excel(writer, sheet_name='buses', index=False)
            pd.DataFrame(simple_generators).to_excel(writer, sheet_name='generators', index=False)
            pd.DataFrame(simple_loads).to_excel(writer, sheet_name='loads', index=False)
            pd.DataFrame(simple_lines).to_excel(writer, sheet_name='lines', index=False)
            pd.DataFrame(simple_links).to_excel(writer, sheet_name='links', index=False)
            pd.DataFrame().to_excel(writer, sheet_name='stores', index=False)  # 빈 저장장치 시트
        
        print(f"\n단순화된 모델이 '{simple_model_file}' 파일에 저장되었습니다.")
        print("이제 아래 명령으로 실행해 보세요:")
        print("python PyPSA_GUI.py --input simplified_input_data.xlsx")
        
        # config_system.py 파일이 존재하면 수정하여 제약조건 제거
        config_file = 'config_system.py'
        if os.path.exists(config_file):
            with open(config_file, 'r', encoding='utf-8') as f:
                config_lines = f.readlines()
            
            # 제약조건 비활성화
            with open(config_file, 'w', encoding='utf-8') as f:
                for line in config_lines:
                    if "constraints" in line and "=" in line:
                        f.write("constraints = {}\n")
                    else:
                        f.write(line)
            
            print(f"\n{config_file} 파일에서 제약조건이 제거되었습니다.")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        
        return False

if __name__ == "__main__":
    simplify_model() 