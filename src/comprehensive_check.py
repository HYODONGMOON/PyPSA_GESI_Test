import pandas as pd
import numpy as np

def comprehensive_check():
    """네트워크의 모든 구성 요소를 종합적으로 확인하는 스크립트"""
    
    # 파일 로드
    input_file = 'integrated_input_data_backup_20250519_134702.xlsx'
    print(f"파일 분석: {input_file}")
    
    # 각 시트 로드
    buses_df = pd.read_excel(input_file, sheet_name='buses')
    lines_df = pd.read_excel(input_file, sheet_name='lines')
    loads_df = pd.read_excel(input_file, sheet_name='loads')
    generators_df = pd.read_excel(input_file, sheet_name='generators')
    links_df = pd.read_excel(input_file, sheet_name='links')
    stores_df = pd.read_excel(input_file, sheet_name='stores')
    
    print(f"버스 수: {len(buses_df)}")
    print(f"라인 수: {len(lines_df)}")
    print(f"로드 수: {len(loads_df)}")
    print(f"발전기 수: {len(generators_df)}")
    print(f"링크 수: {len(links_df)}")
    print(f"저장소 수: {len(stores_df)}")
    
    # 버스 이름 집합
    bus_names = set(buses_df['name'])
    
    # 각 버스의 연결 정보 딕셔너리 초기화
    bus_connections = {bus: {
        'loads': [], 'generators': [], 
        'incoming_lines': [], 'outgoing_lines': [],
        'incoming_links': [], 'outgoing_links': [],
        'stores': []
    } for bus in bus_names}
    
    # 로드 정보 매핑
    for _, load in loads_df.iterrows():
        bus = load['bus']
        if bus in bus_connections:
            bus_connections[bus]['loads'].append(load['name'])
    
    # 발전기 정보 매핑
    for _, gen in generators_df.iterrows():
        bus = gen['bus']
        if bus in bus_connections:
            bus_connections[bus]['generators'].append(gen['name'])
    
    # 라인 정보 매핑
    for _, line in lines_df.iterrows():
        bus0 = line['bus0']
        bus1 = line['bus1']
        
        if bus0 in bus_connections:
            bus_connections[bus0]['outgoing_lines'].append(line['name'])
        if bus1 in bus_connections:
            bus_connections[bus1]['incoming_lines'].append(line['name'])
    
    # 링크 정보 매핑
    for _, link in links_df.iterrows():
        bus0 = link['bus0']
        bus1 = link['bus1']
        
        if bus0 in bus_connections:
            bus_connections[bus0]['outgoing_links'].append(link['name'])
        if bus1 in bus_connections:
            bus_connections[bus1]['incoming_links'].append(link['name'])
    
    # 저장소 정보 매핑
    for _, store in stores_df.iterrows():
        bus = store['bus']
        if bus in bus_connections:
            bus_connections[bus]['stores'].append(store['name'])
    
    # 문제가 있는 버스 분석
    print("\n문제가 있는 버스 분석:")
    problematic_buses = []
    
    for bus, connections in bus_connections.items():
        has_load = len(connections['loads']) > 0
        has_generator = len(connections['generators']) > 0
        has_incoming_line = len(connections['incoming_lines']) > 0
        has_outgoing_line = len(connections['outgoing_lines']) > 0
        has_incoming_link = len(connections['incoming_links']) > 0
        has_outgoing_link = len(connections['outgoing_links']) > 0
        has_store = len(connections['stores']) > 0
        
        # 로드가 있지만 들어오는 전력(발전기, 들어오는 라인, 들어오는 링크)이 없는 경우
        if has_load and not (has_generator or has_incoming_line or has_incoming_link):
            problematic_buses.append({
                'name': bus,
                'connections': connections,
                'reason': "로드가 있지만 발전기, 들어오는 라인, 또는 들어오는 링크가 없습니다."
            })
        
        # 저장소가 있지만 들어오는 전력이 없는 경우
        elif has_store and not (has_generator or has_incoming_line or has_incoming_link or has_load):
            problematic_buses.append({
                'name': bus,
                'connections': connections,
                'reason': "저장소가 있지만 발전기, 들어오는 라인, 또는 들어오는 링크가 없습니다."
            })
        
        # 완전히 고립된 버스
        elif not (has_load or has_generator or has_incoming_line or has_outgoing_line or 
                 has_incoming_link or has_outgoing_link or has_store):
            problematic_buses.append({
                'name': bus,
                'connections': connections,
                'reason': "버스가 완전히 고립되어 있습니다."
            })
    
    # 문제가 있는 버스 출력
    if problematic_buses:
        print(f"  문제가 있는 버스 수: {len(problematic_buses)}")
        for bus_info in problematic_buses:
            bus = bus_info['name']
            reason = bus_info['reason']
            connections = bus_info['connections']
            
            print(f"\n  버스: {bus}")
            print(f"  문제: {reason}")
            print("  연결 정보:")
            print(f"    - 로드: {connections['loads'] if connections['loads'] else '없음'}")
            print(f"    - 발전기: {connections['generators'] if connections['generators'] else '없음'}")
            print(f"    - 들어오는 라인: {connections['incoming_lines'] if connections['incoming_lines'] else '없음'}")
            print(f"    - 나가는 라인: {connections['outgoing_lines'] if connections['outgoing_lines'] else '없음'}")
            print(f"    - 들어오는 링크: {connections['incoming_links'] if connections['incoming_links'] else '없음'}")
            print(f"    - 나가는 링크: {connections['outgoing_links'] if connections['outgoing_links'] else '없음'}")
            print(f"    - 저장소: {connections['stores'] if connections['stores'] else '없음'}")
    else:
        print("  모든 버스가 정상적으로 연결되어 있습니다.")
    
    # 문제 해결 방안 제시
    if problematic_buses:
        print("\n문제 해결 방안:")
        
        # H2와 H 계열 버스가 문제인 경우를 확인
        h2_buses = [bus_info['name'] for bus_info in problematic_buses if '_H2' in bus_info['name']]
        h_buses = [bus_info['name'] for bus_info in problematic_buses if '_H' in bus_info['name'] and not '_H2' in bus_info['name']]
        
        if h2_buses:
            print("\n  H2 버스 문제 해결 방안:")
            print("  1. 각 H2 버스에 Electrolyser 링크를 연결하여 전기에서 수소로 변환 경로를 만들어야 합니다.")
            print("  2. 각 지역의 EL 버스에서 H2 버스로 Electrolyser 링크를 추가합니다.")
            print("  3. 또는 H2 버스에 직접 수소 발전기(예: 'hydrogen')를 추가할 수 있습니다.")
            
        if h_buses:
            print("\n  H 버스 문제 해결 방안:")
            print("  1. 각 H 버스에 Heat Pump 링크를 연결하여 전기에서 열로 변환 경로를 만들어야 합니다.")
            print("  2. 각 지역의 EL 버스에서 H 버스로 Heat Pump 링크를 추가합니다.")
            print("  3. 또는 H 버스에 직접 열 발전기(예: 'heat')를 추가할 수 있습니다.")
        
        print("\n  문제 버스에 저장소가 있는 경우, 연결된 저장소가 충전될 수 있도록 적절한 링크 연결 필요합니다.")
        print("  문제 버스의 로드를 충족시키려면 발전기 또는 다른 버스로부터의 입력이 필요합니다.")

if __name__ == "__main__":
    comprehensive_check() 