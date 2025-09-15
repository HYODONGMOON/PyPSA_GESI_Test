import pandas as pd
from datetime import datetime

def add_heat_hydrogen_slack():
    """열과 수소 버스에 슬랙 발전기를 추가하는 함수"""
    
    print("=== 열/수소 버스 슬랙 발전기 추가 ===")
    
    try:
        # 백업 생성
        import shutil
        backup_file = f"integrated_input_data_before_heat_slack_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        shutil.copy2('integrated_input_data.xlsx', backup_file)
        print(f"백업 파일 생성: {backup_file}")
        
        # 데이터 로드
        input_data = {}
        with pd.ExcelFile('integrated_input_data.xlsx') as xls:
            for sheet in xls.sheet_names:
                input_data[sheet] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet)
        
        generators_df = input_data['generators'].copy()
        buses_df = input_data['buses'].copy()
        loads_df = input_data['loads']
        
        # 모든 버스 유형 식별
        all_buses = set()
        for _, gen in generators_df.iterrows():
            all_buses.add(gen.get('bus', ''))
        for _, load in loads_df.iterrows():
            all_buses.add(load.get('bus', ''))
        for _, bus in buses_df.iterrows():
            all_buses.add(bus.get('name', ''))
        
        # 열 버스와 수소 버스 식별
        heat_buses = [b for b in all_buses if b.endswith('_H') and b in buses_df['name'].values]
        h2_buses = [b for b in all_buses if b.endswith('_H2') and b in buses_df['name'].values]
        
        print(f"열 버스 {len(heat_buses)}개 발견")
        print(f"수소 버스 {len(h2_buses)}개 발견")
        
        new_slack_generators = []
        
        # 열 버스에 슬랙 발전기 추가
        for bus in heat_buses:
            slack_name = f"{bus.replace('_H', '')}_Heat_Slack"
            # 이미 존재하는지 확인
            if slack_name not in generators_df['name'].values:
                slack_gen = {
                    'name': slack_name,
                    'region': bus.split('_')[0],
                    'bus': bus,
                    'carrier': 'heat',
                    'p_nom': 0.0,
                    'p_nom_extendable': True,
                    'p_nom_min': 0.0,
                    'p_nom_max': 50000.0,
                    'marginal_cost': 1e9,  # 매우 높은 비용
                    'capital_cost': 0.0,
                    'efficiency': 1.0,
                    'p_max_pu': 1.0,
                    'committable': False
                }
                new_slack_generators.append(slack_gen)
                print(f"열 슬랙 추가: {slack_name} -> {bus}")
        
        # 수소 버스에 슬랙 발전기 추가
        for bus in h2_buses:
            slack_name = f"{bus.replace('_H2', '')}_H2_Slack"
            # 이미 존재하는지 확인
            if slack_name not in generators_df['name'].values:
                slack_gen = {
                    'name': slack_name,
                    'region': bus.split('_')[0],
                    'bus': bus,
                    'carrier': 'hydrogen',
                    'p_nom': 0.0,
                    'p_nom_extendable': True,
                    'p_nom_min': 0.0,
                    'p_nom_max': 50000.0,
                    'marginal_cost': 1e9,  # 매우 높은 비용
                    'capital_cost': 0.0,
                    'efficiency': 1.0,
                    'p_max_pu': 1.0,
                    'committable': False
                }
                new_slack_generators.append(slack_gen)
                print(f"수소 슬랙 추가: {slack_name} -> {bus}")
        
        # 새 슬랙 발전기 추가
        if new_slack_generators:
            new_slack_df = pd.DataFrame(new_slack_generators)
            generators_df = pd.concat([generators_df, new_slack_df], ignore_index=True)
            print(f"\n총 {len(new_slack_generators)}개 슬랙 발전기 추가됨")
        else:
            print("추가할 슬랙 발전기가 없습니다")
        
        # 수정된 파일 저장
        with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
            generators_df.to_excel(writer, sheet_name='generators', index=False)
            for sheet_name, df in input_data.items():
                if sheet_name != 'generators':
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n=== 완료 ===")
        print(f"- 열 버스 슬랙: {len([g for g in new_slack_generators if 'Heat_Slack' in g['name']])}개 추가")
        print(f"- 수소 버스 슬랙: {len([g for g in new_slack_generators if 'H2_Slack' in g['name']])}개 추가")
        print("- 이제 CHP의 열 생산과 전해조의 수소 생산이 infeasible을 일으키지 않을 것입니다")
        
        # 최종 확인
        final_gen_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
        total_slack = len(final_gen_df[final_gen_df['name'].str.contains('Slack', na=False)])
        print(f"- 총 슬랙 발전기: {total_slack}개 (전력 17개 + 열/수소 {len(new_slack_generators)}개)")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    add_heat_hydrogen_slack() 