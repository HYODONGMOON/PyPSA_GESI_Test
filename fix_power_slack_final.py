import pandas as pd
from datetime import datetime
import shutil

def fix_power_slack():
    """전력 슬랙 발전기가 제대로 작동하도록 수정하는 함수"""
    
    print("=== 전력 슬랙 발전기 최종 수정 ===")
    
    try:
        # 백업 생성
        backup_file = f"integrated_input_data_before_power_slack_fix_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        shutil.copy2('integrated_input_data.xlsx', backup_file)
        print(f"백업 파일 생성: {backup_file}")
        
        # 데이터 로드
        input_data = {}
        with pd.ExcelFile('integrated_input_data.xlsx') as xls:
            for sheet in xls.sheet_names:
                input_data[sheet] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet)
        
        generators_df = input_data['generators'].copy()
        buses_df = input_data['buses'].copy()
        
        # 현재 전력 슬랙 발전기 확인
        current_power_slack = generators_df[
            generators_df['name'].str.contains('Slack_Gen', na=False) & 
            generators_df['bus'].str.contains('_EL', na=False)
        ]
        
        print(f"현재 전력 슬랙 발전기 개수: {len(current_power_slack)}")
        
        # 전력 버스 목록 확인 (carrier가 AC로 되어 있음)
        el_buses = buses_df[buses_df['carrier'] == 'AC']['name'].unique()
        print(f"전력 버스 개수: {len(el_buses)}")
        
        # 기존 전력 슬랙 발전기 제거 (다시 추가하기 위해)
        generators_df = generators_df[~(
            generators_df['name'].str.contains('Slack_Gen', na=False) & 
            generators_df['bus'].str.contains('_EL', na=False)
        )]
        
        print("기존 전력 슬랙 발전기 제거됨")
        
        # 모든 전력 버스에 새로운 슬랙 발전기 추가
        new_slack_generators = []
        
        for bus in el_buses:
            slack_name = f"{bus.replace('_EL', '')}_Power_Slack_Critical"
            
            slack_gen = {
                'name': slack_name,
                'region': bus.replace('_EL', ''),
                'bus': bus,
                'carrier': 'electricity',
                'p_nom': 0.0,
                'p_nom_extendable': True,
                'p_nom_min': 0.0,
                'p_nom_max': 1000000.0,  # 매우 큰 값으로 설정
                'marginal_cost': 1e9,    # 매우 높은 비용
                'capital_cost': 0.0,
                'efficiency': 1.0,
                'p_max_pu': 1.0,
                'committable': False
            }
            new_slack_generators.append(slack_gen)
            print(f"전력 슬랙 추가: {slack_name} -> {bus}")
        
        # 새 슬랙 발전기를 데이터프레임에 추가
        if new_slack_generators:
            new_slack_df = pd.DataFrame(new_slack_generators)
            generators_df = pd.concat([generators_df, new_slack_df], ignore_index=True)
            print(f"새로운 전력 슬랙 발전기 {len(new_slack_generators)}개 추가됨")
        
        # 업데이트된 데이터 저장
        input_data['generators'] = generators_df
        
        with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
            for sheet_name, df in input_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print("\n=== 전력 슬랙 발전기 수정 완료 ===")
        print(f"총 {len(new_slack_generators)}개의 전력 슬랙 발전기가 추가되었습니다.")
        print("이제 PyPSA 분석을 다시 실행하면 전력 슬랙이 작동할 것입니다.")
        
        # 검증
        print("\n=== 검증 ===")
        verification_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
        verification_slack = verification_df[
            verification_df['name'].str.contains('Power_Slack_Critical', na=False)
        ]
        print(f"검증: 전력 슬랙 발전기 {len(verification_slack)}개 확인됨")
        
    except Exception as e:
        print(f"오류 발생: {e}")
        return False
        
    return True

if __name__ == "__main__":
    fix_power_slack() 