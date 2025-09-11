import pandas as pd

def check_network_problems():
    """네트워크 문제점 검사를 위한 스크립트"""
    
    # 파일 로드
    input_file = 'integrated_input_data_backup_20250519_134702.xlsx'
    print(f"파일 분석: {input_file}")
    
    # 각 시트 로드
    buses_df = pd.read_excel(input_file, sheet_name='buses')
    lines_df = pd.read_excel(input_file, sheet_name='lines')
    loads_df = pd.read_excel(input_file, sheet_name='loads')
    generators_df = pd.read_excel(input_file, sheet_name='generators')
    
    print(f"버스 수: {len(buses_df)}")
    print(f"라인 수: {len(lines_df)}")
    print(f"로드 수: {len(loads_df)}")
    print(f"발전기 수: {len(generators_df)}")
    
    # 1. BSN_GND 라인 검사
    print("\n1. BSN_GND 라인 검사:")
    bsn_gnd_line = lines_df[lines_df['name'] == 'BSN_GND']
    if len(bsn_gnd_line) > 0:
        print("  BSN_GND 라인 정보:")
        for _, line in bsn_gnd_line.iterrows():
            print(f"    - 이름: {line['name']}")
            print(f"    - bus0: {line['bus0']}")
            print(f"    - bus1: {line['bus1']}")
            
            # bus0와 bus1이 버스 시트에 존재하는지 확인
            bus0_exists = line['bus0'] in buses_df['name'].values
            bus1_exists = line['bus1'] in buses_df['name'].values
            
            print(f"    - bus0 존재: {bus0_exists}")
            print(f"    - bus1 존재: {bus1_exists}")
            
            # 문제점 표시
            if not bus0_exists:
                print(f"    [문제] bus0 '{line['bus0']}'가 버스 시트에 정의되지 않았습니다.")
            if not bus1_exists:
                print(f"    [문제] bus1 '{line['bus1']}'가 버스 시트에 정의되지 않았습니다.")
    else:
        print("  BSN_GND 라인을 찾을 수 없습니다.")
    
    # 2. 모든 라인의 bus0/bus1 존재 여부 확인
    print("\n2. 모든 라인의 버스 존재 여부 확인:")
    bus_names = set(buses_df['name'])
    
    problematic_lines = []
    for idx, line in lines_df.iterrows():
        bus0 = line['bus0']
        bus1 = line['bus1']
        
        if bus0 not in bus_names or bus1 not in bus_names:
            problematic_lines.append({
                'name': line['name'],
                'bus0': bus0,
                'bus0_exists': bus0 in bus_names,
                'bus1': bus1,
                'bus1_exists': bus1 in bus_names
            })
    
    if problematic_lines:
        print(f"  문제가 있는 라인 수: {len(problematic_lines)}")
        for line in problematic_lines:
            print(f"  - 라인: {line['name']}")
            if not line['bus0_exists']:
                print(f"    [문제] bus0 '{line['bus0']}'가 정의되지 않았습니다.")
            if not line['bus1_exists']:
                print(f"    [문제] bus1 '{line['bus1']}'가 정의되지 않았습니다.")
    else:
        print("  모든 라인의 버스가 제대로 정의되어 있습니다.")
    
    # 3. 로드가 있지만 들어오는 연결이 없는 버스 확인
    print("\n3. 로드가 있지만 들어오는 연결이 없는 버스 확인:")
    
    # 각 버스의 정보 구성
    bus_info = {bus: {'loads': [], 'generators': [], 'incoming_lines': [], 'outgoing_lines': []} 
                for bus in bus_names}
    
    # 로드 정보 매핑
    for _, load in loads_df.iterrows():
        bus = load['bus']
        if bus in bus_info:
            bus_info[bus]['loads'].append(load['name'])
    
    # 발전기 정보 매핑
    for _, gen in generators_df.iterrows():
        bus = gen['bus']
        if bus in bus_info:
            bus_info[bus]['generators'].append(gen['name'])
    
    # 라인 정보 매핑
    for _, line in lines_df.iterrows():
        bus0 = line['bus0']
        bus1 = line['bus1']
        
        if bus0 in bus_info:
            bus_info[bus0]['outgoing_lines'].append(line['name'])
        if bus1 in bus_info:
            bus_info[bus1]['incoming_lines'].append(line['name'])
    
    # 문제 있는 버스 확인
    problematic_buses = []
    for bus, info in bus_info.items():
        has_load = len(info['loads']) > 0
        has_generator = len(info['generators']) > 0
        has_incoming = len(info['incoming_lines']) > 0
        
        # 로드가 있지만 들어오는 연결 또는 발전기가 없는 경우
        if has_load and not (has_generator or has_incoming):
            problematic_buses.append({
                'name': bus,
                'loads': info['loads'],
                'generators': info['generators'],
                'incoming_lines': info['incoming_lines'],
                'outgoing_lines': info['outgoing_lines']
            })
    
    if problematic_buses:
        print(f"  문제가 있는 버스 수: {len(problematic_buses)}")
        for bus in problematic_buses:
            print(f"  - 버스: {bus['name']}")
            print(f"    - 로드: {bus['loads']}")
            print(f"    - 발전기: {bus['generators']}")
            print(f"    - 들어오는 라인: {bus['incoming_lines']}")
            print(f"    - 나가는 라인: {bus['outgoing_lines']}")
            print(f"    [문제] 로드가 있지만 발전기 또는 들어오는 연결이 없습니다.")
    else:
        print("  모든 로드가 있는 버스에는 적절한 전력 공급이 있습니다.")

if __name__ == "__main__":
    check_network_problems() 