import pandas as pd
import numpy as np

def fix_line_connections():
    """선로 연결 문제 해결"""
    print("=== 선로 연결 문제 해결 시작 ===")
    
    # 백업 파일 생성
    from datetime import datetime
    backup_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f'integrated_input_data_backup_{backup_time}.xlsx'
    
    # 원본 파일 백업
    import shutil
    shutil.copy('integrated_input_data.xlsx', backup_filename)
    print(f"백업 파일 생성: {backup_filename}")
    
    # 데이터 로드
    input_data = {}
    xls = pd.ExcelFile('integrated_input_data.xlsx')
    for sheet_name in xls.sheet_names:
        input_data[sheet_name] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet_name)
    
    # 현재 버스 목록 확인
    buses_df = input_data['buses']
    print(f"\n현재 버스 수: {len(buses_df)}")
    print("버스 이름 예시:")
    for i, bus_name in enumerate(buses_df['name'][:10]):
        print(f"  {i+1}. {bus_name}")
    
    # 선로 데이터 확인
    lines_df = input_data['lines']
    print(f"\n현재 선로 수: {len(lines_df)}")
    print("선로 연결 예시:")
    for i, (_, line) in enumerate(lines_df.head(5).iterrows()):
        print(f"  {i+1}. {line['name']}: {line['bus0']} -> {line['bus1']}")
    
    # 버스 매핑 생성 (지역 코드 -> 실제 전력 버스)
    bus_mapping = {}
    for _, bus in buses_df.iterrows():
        bus_name = str(bus['name'])
        if '_EL' in bus_name:  # 전력 버스만 매핑
            region = bus_name.split('_')[0]
            bus_mapping[region] = bus_name
    
    print(f"\n버스 매핑:")
    for region, bus_name in sorted(bus_mapping.items()):
        print(f"  {region} -> {bus_name}")
    
    # 선로 버스 이름 수정
    print(f"\n선로 버스 이름 수정 중...")
    modified_lines = 0
    
    for idx, line in lines_df.iterrows():
        bus0_original = str(line['bus0'])
        bus1_original = str(line['bus1'])
        
        # bus0 매핑
        if bus0_original in bus_mapping:
            lines_df.at[idx, 'bus0'] = bus_mapping[bus0_original]
            print(f"  {line['name']}: bus0 {bus0_original} -> {bus_mapping[bus0_original]}")
            modified_lines += 1
        
        # bus1 매핑
        if bus1_original in bus_mapping:
            lines_df.at[idx, 'bus1'] = bus_mapping[bus1_original]
            print(f"  {line['name']}: bus1 {bus1_original} -> {bus_mapping[bus1_original]}")
    
    print(f"\n수정된 선로 수: {modified_lines}")
    
    # 수정된 선로 연결 확인
    print(f"\n수정된 선로 연결:")
    valid_connections = 0
    invalid_connections = 0
    
    for _, line in lines_df.iterrows():
        bus0 = str(line['bus0'])
        bus1 = str(line['bus1'])
        
        # 버스 존재 확인
        bus0_exists = bus0 in buses_df['name'].values
        bus1_exists = bus1 in buses_df['name'].values
        
        if bus0_exists and bus1_exists:
            print(f"  ✓ {line['name']}: {bus0} -> {bus1}")
            valid_connections += 1
        else:
            print(f"  ✗ {line['name']}: {bus0} -> {bus1} (버스 없음)")
            invalid_connections += 1
    
    print(f"\n연결 상태:")
    print(f"  유효한 연결: {valid_connections}개")
    print(f"  무효한 연결: {invalid_connections}개")
    
    # 수정된 데이터 저장
    input_data['lines'] = lines_df
    
    with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
        for sheet_name, df in input_data.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"\n수정된 데이터가 저장되었습니다.")
    
    return valid_connections, invalid_connections

if __name__ == "__main__":
    fix_line_connections() 