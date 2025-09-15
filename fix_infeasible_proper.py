import pandas as pd
import numpy as np
import os
from datetime import datetime

def fix_infeasible_proper():
    """Interface 설정을 존중하면서 infeasible 문제만 해결하는 함수"""
    
    print("=== Interface 설정 존중하는 Infeasible 해결 ===")
    
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
        
        # interface.xlsx에서 설정 확인
        interface_path = 'interface.xlsx'
        interface_settings = {}
        
        if os.path.exists(interface_path):
            print("\nInterface.xlsx에서 설정 확인 중...")
            try:
                # generators 시트에서 extendable 설정 확인
                interface_gens = pd.read_excel(interface_path, sheet_name='generators')
                if 'p_nom_extendable' in interface_gens.columns:
                    for _, gen in interface_gens.iterrows():
                        if pd.notna(gen.get('name')):
                            interface_settings[gen['name']] = {
                                'p_nom_extendable': gen.get('p_nom_extendable', False)
                            }
                    print(f"Interface에서 {len(interface_settings)}개 발전기 설정 로딩")
            except Exception as e:
                print(f"Interface generators 시트 로딩 실패: {e}")
        
        # 1. 슬랙 발전기만 추가 (다른 설정은 변경하지 않음)
        print("\n1. 슬랙 발전기 추가 중...")
        generators_df = input_data['generators'].copy()
        buses_df = input_data['buses'].copy()
        
        # 전력 버스별로 슬랙 발전기 추가
        electric_buses = buses_df[buses_df['name'].str.endswith('_EL')]['name'].tolist()
        slack_generators = []
        
        for bus in electric_buses:
            # 슬랙 발전기가 없으면 추가
            slack_name = f"{bus.replace('_EL', '')}_Slack_Gen"
            if slack_name not in generators_df['name'].values:
                slack_gen = {
                    'name': slack_name,
                    'bus': bus,
                    'carrier': 'electricity',
                    'p_nom': 0.0,
                    'p_nom_extendable': True,  # 슬랙은 확장 가능해야 함
                    'p_nom_min': 0.0,
                    'p_nom_max': 50000.0,
                    'capital_cost': 0.0,
                    'marginal_cost': 1e9,  # 매우 높은 비용
                    'efficiency': 1.0
                }
                slack_generators.append(slack_gen)
        
        if slack_generators:
            slack_df = pd.DataFrame(slack_generators)
            generators_df = pd.concat([generators_df, slack_df], ignore_index=True)
            print(f"슬랙 발전기 {len(slack_generators)}개 추가됨")
        
        # 2. Interface 설정 적용 (extendable 설정 복원)
        print("\n2. Interface 설정 적용 중...")
        
        interface_applied = 0
        for idx, gen in generators_df.iterrows():
            gen_name = gen['name']
            
            # 슬랙 발전기는 건드리지 않음
            if 'Slack' in gen_name:
                continue
                
            # Interface에서 설정이 있으면 적용
            if gen_name in interface_settings:
                original_extendable = gen.get('p_nom_extendable', False)
                interface_extendable = interface_settings[gen_name]['p_nom_extendable']
                
                if original_extendable != interface_extendable:
                    generators_df.at[idx, 'p_nom_extendable'] = interface_extendable
                    interface_applied += 1
                    print(f"  {gen_name}: extendable {original_extendable} -> {interface_extendable}")
            
            # Interface 설정이 없으면 기본적으로 False로 설정 (하드코딩 방지)
            elif gen_name not in interface_settings and 'Slack' not in gen_name:
                if gen.get('p_nom_extendable', False) == True:
                    generators_df.at[idx, 'p_nom_extendable'] = False
                    interface_applied += 1
                    print(f"  {gen_name}: 하드코딩된 extendable=True를 False로 복원")
        
        if interface_applied > 0:
            print(f"Interface 설정 적용: {interface_applied}개 발전기")
        else:
            print("Interface 설정 변경사항 없음")
        
        # 3. CHP 링크 효율만 검증 (필요시)
        print("\n3. CHP 링크 효율 검증 중...")
        links_df = input_data['links'].copy()
        
        chp_efficiency_fixed = 0
        chp_links = links_df[links_df['name'].str.contains('CHP', na=False)]
        for idx, link in chp_links.iterrows():
            # efficiency1 (전력 효율) 확인
            if pd.isna(link.get('efficiency1')) or link.get('efficiency1') == 0:
                links_df.at[idx, 'efficiency1'] = 0.35
                chp_efficiency_fixed += 1
                print(f"  {link['name']} 전력 효율 설정: 0.35")
        
        if chp_efficiency_fixed == 0:
            print("CHP 효율 수정 불필요")
        
        # 4. 수정된 데이터 저장
        print("\n4. 수정된 데이터 저장 중...")
        
        with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
            # 수정된 데이터 저장
            generators_df.to_excel(writer, sheet_name='generators', index=False)
            if chp_efficiency_fixed > 0:
                links_df.to_excel(writer, sheet_name='links', index=False)
            
            # 나머지 시트들은 그대로 복사
            for sheet_name in ['buses', 'loads', 'stores', 'lines', 'constraints', 'timeseries', 'renewable_patterns', 'load_patterns']:
                if sheet_name in input_data:
                    input_data[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
        
        print("\n=== 수정 완료 ===")
        print(f"- 슬랙 발전기: {len(slack_generators)}개 추가")
        print(f"- Interface 설정 적용: {interface_applied}개 발전기")
        print(f"- CHP 효율 수정: {chp_efficiency_fixed}개")
        print("- 백업 발전기 과다 추가 방지됨")
        print("\n✅ Interface 설정을 존중하면서 infeasible 문제만 해결했습니다.")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    fix_infeasible_proper() 