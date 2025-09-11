import pandas as pd
import numpy as np
from datetime import datetime
import shutil

def fix_line_bus_names():
    """선로의 버스 이름을 실제 버스 이름 형식에 맞게 수정"""
    
    print("=== 선로 버스 이름 수정 시작 ===")
    
    # 백업 생성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f"integrated_input_data_backup_{timestamp}.xlsx"
    shutil.copy2("integrated_input_data.xlsx", backup_filename)
    print(f"백업 파일 생성: {backup_filename}")
    
    # 데이터 로드
    try:
        # 모든 시트 읽기
        with pd.ExcelFile('integrated_input_data.xlsx') as xls:
            all_data = {}
            for sheet_name in xls.sheet_names:
                all_data[sheet_name] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet_name)
        
        # 버스 이름 확인
        buses_df = all_data['buses']
        lines_df = all_data['lines']
        
        print(f"\n현재 버스 수: {len(buses_df)}")
        print("버스 이름 예시:")
        for i, bus_name in enumerate(buses_df['name'].head(10)):
            print(f"  {i+1}. {bus_name}")
        
        print(f"\n현재 선로 수: {len(lines_df)}")
        print("선로 연결 예시:")
        for i, (_, line) in enumerate(lines_df.head(5).iterrows()):
            print(f"  {i+1}. {line['name']}: {line['bus0']} -> {line['bus1']}")
        
        # 버스 이름 매핑 생성
        # 실제 버스: BSN_BSN_EL, CBD_CBD_EL, ...
        # 선로에서 참조: BSN, CBD, ...
        
        bus_mapping = {}
        actual_buses = set(buses_df['name'])
        
        # 각 지역의 전력 버스 찾기
        for bus_name in actual_buses:
            if '_EL' in bus_name:  # 전력 버스만
                # BSN_BSN_EL -> BSN
                region_code = bus_name.split('_')[0]
                bus_mapping[region_code] = bus_name
        
        print(f"\n버스 매핑:")
        for old_name, new_name in bus_mapping.items():
            print(f"  {old_name} -> {new_name}")
        
        # 선로의 버스 이름 수정
        modified_lines = 0
        for idx, line in lines_df.iterrows():
            bus0_old = str(line['bus0'])
            bus1_old = str(line['bus1'])
            
            # 버스 이름 변경
            bus0_new = bus_mapping.get(bus0_old, bus0_old)
            bus1_new = bus_mapping.get(bus1_old, bus1_old)
            
            if bus0_new != bus0_old or bus1_new != bus1_old:
                lines_df.at[idx, 'bus0'] = bus0_new
                lines_df.at[idx, 'bus1'] = bus1_new
                modified_lines += 1
                print(f"선로 {line['name']}: {bus0_old}->{bus0_new}, {bus1_old}->{bus1_new}")
        
        print(f"\n수정된 선로 수: {modified_lines}")
        
        # 수정된 데이터 확인
        print(f"\n수정 후 선로 연결 확인:")
        valid_connections = 0
        invalid_connections = 0
        
        for _, line in lines_df.iterrows():
            bus0 = line['bus0']
            bus1 = line['bus1']
            
            if bus0 in actual_buses and bus1 in actual_buses:
                valid_connections += 1
            else:
                invalid_connections += 1
                print(f"  여전히 유효하지 않음: {line['name']} ({bus0} -> {bus1})")
                if bus0 not in actual_buses:
                    print(f"    누락된 버스: {bus0}")
                if bus1 not in actual_buses:
                    print(f"    누락된 버스: {bus1}")
        
        print(f"\n연결 상태:")
        print(f"  유효한 연결: {valid_connections}")
        print(f"  유효하지 않은 연결: {invalid_connections}")
        
        # 파일 저장
        if modified_lines > 0:
            all_data['lines'] = lines_df
            
            # Excel 파일로 저장
            with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
                for sheet_name, df in all_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"\n수정된 데이터가 저장되었습니다.")
            print(f"총 {modified_lines}개 선로의 버스 이름이 수정되었습니다.")
        else:
            print("\n수정할 내용이 없습니다.")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    fix_line_bus_names() 