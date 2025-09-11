import pandas as pd
import os
import shutil
from datetime import datetime

def fix_optimization():
    """
    최적화 문제를 해결하기 위해 네트워크 구성을 추가로 수정하는 함수
    
    1. 외부 그리드에 부하 추가
    2. 링크 제약조건 완화
    3. 부하와 발전 사이의 균형 확인
    """
    print("최적화 문제 해결 중...")
    
    # 파일 경로
    input_file = 'integrated_input_data.xlsx'
    
    # 백업 파일 생성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = f'integrated_input_data_backup_opt_{timestamp}.xlsx'
    
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
            links = pd.read_excel(xls, sheet_name='links')
            generators = pd.read_excel(xls, sheet_name='generators')
            loads = pd.read_excel(xls, sheet_name='loads')
            lines = pd.read_excel(xls, sheet_name='lines')
            stores = pd.read_excel(xls, sheet_name='stores')
        
        # 1. 모든 링크의 확장성과 용량 제약 확인 및 수정
        print("링크 제약조건 완화 중...")
        for idx in links.index:
            # 모든 링크를 확장 가능하게 설정
            links.loc[idx, 'p_nom_extendable'] = True
            
            # 최대 용량 제약 완화
            if 'p_nom_max' in links.columns:
                links.loc[idx, 'p_nom_max'] = 10000  # 매우 큰 값으로 설정
            else:
                links['p_nom_max'] = 10000
            
            # 최소 용량을 0으로 설정
            if 'p_nom_min' in links.columns:
                links.loc[idx, 'p_nom_min'] = 0
            else:
                links['p_nom_min'] = 0
        
        print(f"총 {len(links)}개 링크의 제약조건이 완화되었습니다.")
        
        # 2. 모든 라인의 확장성과 용량 제약 확인 및 수정
        print("라인 제약조건 완화 중...")
        for idx in lines.index:
            # 모든 라인을 확장 가능하게 설정
            if 's_nom_extendable' in lines.columns:
                lines.loc[idx, 's_nom_extendable'] = True
            else:
                lines['s_nom_extendable'] = True
                
            # 최대 용량 제약 완화
            if 's_nom_max' in lines.columns:
                lines.loc[idx, 's_nom_max'] = 10000
            else:
                lines['s_nom_max'] = 10000
                
            # 최소 용량을 0으로 설정
            if 's_nom_min' in lines.columns:
                lines.loc[idx, 's_nom_min'] = 0
            else:
                lines['s_nom_min'] = 0
        
        print(f"총 {len(lines)}개 라인의 제약조건이 완화되었습니다.")
        
        # 3. 외부 그리드에 부하 추가 (수출용)
        if 'External_Grid' in buses['name'].values:
            # 기존 부하 확인
            external_loads = loads[loads['bus'] == 'External_Grid']
            
            # 외부 그리드에 부하가 없으면 추가
            if external_loads.empty:
                # 총 부하량 계산
                total_load = loads['p_set'].sum()
                
                # 외부 그리드에 작은 부하 추가 (총 부하의 0.1%)
                external_load = {
                    'name': 'External_Grid_Load',
                    'bus': 'External_Grid',
                    'p_set': total_load * 0.001,  # 아주 작은 값으로 설정
                    'carrier': 'AC'
                }
                
                loads = pd.concat([loads, pd.DataFrame([external_load])], ignore_index=True)
                print(f"External_Grid에 {external_load['p_set']:.2f} MW의 부하가 추가되었습니다.")
        
        # 4. 각 지역별 부하와 발전 용량 계산 및 표시
        print("\n지역별 부하와 발전 용량:")
        
        region_codes = set()
        
        # 지역 코드 추출
        for bus in buses['name']:
            if '_EL' in bus:
                region_code = bus.split('_')[0]
                region_codes.add(region_code)
        
        # 지역별 분석
        for region in sorted(region_codes):
            region_buses = buses[buses['name'].str.startswith(f"{region}_")]
            region_bus_names = region_buses['name'].tolist()
            
            # 지역 버스에 연결된 부하 계산
            region_loads = loads[loads['bus'].isin(region_bus_names)]
            total_region_load = region_loads['p_set'].sum()
            
            # 지역 버스에 연결된 발전 용량 계산
            region_gens = generators[generators['bus'].isin(region_bus_names)]
            total_region_gen_capacity = region_gens['p_nom'].sum()
            
            # 확장 가능한 발전 용량 계산
            extendable_gens = region_gens[region_gens['p_nom_extendable'] == True]
            extendable_capacity = extendable_gens['p_nom'].sum()
            max_extendable = 0
            if 'p_nom_max' in extendable_gens.columns:
                max_extendable = extendable_gens['p_nom_max'].sum()
            
            print(f"{region}: 부하 {total_region_load:.2f} MW, 발전 용량 {total_region_gen_capacity:.2f} MW, 확장 가능 {extendable_capacity:.2f}/{max_extendable:.2f} MW")
        
        # 5. 모든 발전기의 제약조건 완화
        print("\n발전기 제약조건 완화 중...")
        for idx in generators.index:
            # 발전기를 확장 가능하게 설정
            generators.loc[idx, 'p_nom_extendable'] = True
            
            # 최대 용량 제약 완화
            if 'p_nom_max' in generators.columns:
                if pd.notna(generators.loc[idx, 'p_nom_max']):
                    generators.loc[idx, 'p_nom_max'] = generators.loc[idx, 'p_nom_max'] * 5  # 기존 제약의 5배로 설정
                else:
                    generators.loc[idx, 'p_nom_max'] = 10000
            else:
                generators['p_nom_max'] = 10000
        
        print(f"총 {len(generators)}개 발전기의 제약조건이 완화되었습니다.")
        
        # 6. slack 버스 추가 (더미 발전기와 부하로)
        if 'Slack_Bus' not in buses['name'].values:
            # Slack 버스 추가
            slack_bus = {
                'name': 'Slack_Bus',
                'v_nom': 345,
                'carrier': 'AC',
                'x': 0,  # 임의의 좌표
                'y': 0   # 임의의 좌표
            }
            buses = pd.concat([buses, pd.DataFrame([slack_bus])], ignore_index=True)
            
            # Slack 발전기 추가
            slack_gen = {
                'name': 'Slack_Generator',
                'bus': 'Slack_Bus',
                'carrier': 'AC',
                'p_nom': 100000,
                'p_nom_extendable': True,
                'marginal_cost': 10000,  # 아주 높은 비용
                'p_nom_min': 0,
                'p_nom_max': 100000
            }
            generators = pd.concat([generators, pd.DataFrame([slack_gen])], ignore_index=True)
            
            # Slack 부하 추가
            slack_load = {
                'name': 'Slack_Load',
                'bus': 'Slack_Bus',
                'p_set': 0.1,  # 아주 작은 값
                'carrier': 'AC'
            }
            loads = pd.concat([loads, pd.DataFrame([slack_load])], ignore_index=True)
            
            # External_Grid와 Slack_Bus 연결
            slack_link = {
                'name': 'External_To_Slack',
                'bus0': 'External_Grid',
                'bus1': 'Slack_Bus',
                'p_nom': 100000,
                'p_nom_extendable': True,
                'p_nom_min': 0,
                'p_nom_max': 100000,
                'efficiency': 1.0,
                'marginal_cost': 0.1
            }
            links = pd.concat([links, pd.DataFrame([slack_link])], ignore_index=True)
            
            print("Slack 버스, 발전기, 부하 및 연결 링크가 추가되었습니다.")
        
        # 엑셀 파일에 저장
        with pd.ExcelWriter(input_file) as writer:
            buses.to_excel(writer, sheet_name='buses', index=False)
            links.to_excel(writer, sheet_name='links', index=False)
            generators.to_excel(writer, sheet_name='generators', index=False)
            loads.to_excel(writer, sheet_name='loads', index=False)
            lines.to_excel(writer, sheet_name='lines', index=False)
            stores.to_excel(writer, sheet_name='stores', index=False)
        
        print(f"\n{input_file} 파일이 성공적으로 수정되었습니다.")
        print("이제 PyPSA_GUI.py를 다시 실행해보세요.")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        
        # 백업에서 복원
        print(f"백업에서 복원 중...")
        shutil.copy2(backup_file, input_file)
        print(f"원본 파일이 백업에서 복원되었습니다.")
        
        return False

if __name__ == "__main__":
    fix_optimization() 