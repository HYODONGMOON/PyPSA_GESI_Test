import pandas as pd
import os

def debug_power_slack():
    """전력 슬랙 발전기 추가 로직 디버그"""
    
    print("=== 전력 슬랙 발전기 디버그 ===")
    
    try:
        # 환경변수 확인
        disable_slack = os.environ.get('DISABLE_POWER_SLACK', '0')
        print(f"DISABLE_POWER_SLACK 환경변수: {disable_slack}")
        print(f"슬랙 활성화 조건: {disable_slack} != '1' = {disable_slack != '1'}")
        
        # integrated_input_data에서 버스 정보 확인
        buses_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='buses')
        print(f"\n버스 개수: {len(buses_df)}")
        print(f"버스 컬럼: {list(buses_df.columns)}")
        
        # 전력 버스 확인
        if 'carrier' in buses_df.columns:
            print(f"\n캐리어 유형들: {buses_df['carrier'].value_counts().to_dict()}")
            
            # AC 캐리어 버스들
            ac_buses = buses_df[buses_df['carrier'] == 'AC']
            print(f"\nAC 캐리어 버스 개수: {len(ac_buses)}")
            
            if len(ac_buses) > 0:
                print("처음 5개 AC 버스:")
                for _, bus in ac_buses.head().iterrows():
                    bus_name = bus['name']
                    carrier = bus['carrier']
                    print(f"- 버스: {bus_name}, 캐리어: '{carrier}'")
                    
                    # 조건 테스트
                    condition_result = (carrier == 'electricity') or (carrier == 'AC')
                    print(f"  조건 테스트: ('{carrier}' == 'electricity') or ('{carrier}' == 'AC') = {condition_result}")
        
        # 기존 슬랙 발전기 확인
        gen_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
        el_slack = gen_df[gen_df['name'].str.contains('Slack', na=False) & gen_df['bus'].str.contains('_EL', na=False)]
        print(f"\n입력 데이터의 전력 슬랙 발전기: {len(el_slack)}개")
        
    except Exception as e:
        print(f"오류: {e}")

if __name__ == "__main__":
    debug_power_slack() 