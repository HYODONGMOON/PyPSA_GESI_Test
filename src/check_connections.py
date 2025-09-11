import pandas as pd
import os
import shutil
from datetime import datetime

def check_connections():
    """현재 버스와 선로 연결 상태 확인"""
    
    try:
        # 데이터 로드
        buses_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='buses')
        lines_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='lines')
        
        print("=== 연결 상태 확인 ===")
        
        # 버스 정보
        print(f"\n버스 수: {len(buses_df)}")
        print("첫 5개 버스:")
        for i, bus_name in enumerate(buses_df['name'].head(5)):
            print(f"  {i+1}. {bus_name}")
        
        # 선로 정보
        print(f"\n선로 수: {len(lines_df)}")
        print("첫 5개 선로 연결:")
        for i, (_, line) in enumerate(lines_df.head(5).iterrows()):
            print(f"  {i+1}. {line['name']}: {line['bus0']} -> {line['bus1']}")
        
        # 연결 검사
        actual_buses = set(buses_df['name'])
        valid_count = 0
        invalid_count = 0
        
        print(f"\n선로 연결 검사:")
        for _, line in lines_df.iterrows():
            bus0 = str(line['bus0'])
            bus1 = str(line['bus1'])
            
            if bus0 in actual_buses and bus1 in actual_buses:
                valid_count += 1
            else:
                invalid_count += 1
                print(f"  유효하지 않음: {line['name']} ({bus0} -> {bus1})")
        
        print(f"\n결과:")
        print(f"  유효한 연결: {valid_count}")
        print(f"  유효하지 않은 연결: {invalid_count}")
        
        # 전력 버스 매핑
        print(f"\n전력 버스 매핑 확인:")
        el_buses = [bus for bus in actual_buses if '_EL' in bus]
        print(f"전력 버스 수: {len(el_buses)}")
        for bus in el_buses[:5]:
            region = bus.split('_')[0]
            print(f"  {region} -> {bus}")
        
    except Exception as e:
        print(f"오류: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_connections() 