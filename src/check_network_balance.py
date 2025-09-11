import pandas as pd
import numpy as np
from datetime import datetime

def check_network_balance():
    """
    네트워크의 버스, 라인, 로드 사이의 균형을 확인하는 스크립트
    """
    input_file = 'integrated_input_data_backup_20250519_134702.xlsx'
    print(f"파일 분석: {input_file}")
    print(f"분석 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*70)
    
    # 모든 시트 이름 확인
    with pd.ExcelFile(input_file) as xls:
        sheet_names = xls.sheet_names
        print(f"엑셀 파일 시트: {sheet_names}")
    
    # 각 시트 로드
    try:
        buses_df = pd.read_excel(input_file, sheet_name='buses')
        lines_df = pd.read_excel(input_file, sheet_name='lines')
        loads_df = pd.read_excel(input_file, sheet_name='loads')
        generators_df = pd.read_excel(input_file, sheet_name='generators')
        links_df = pd.read_excel(input_file, sheet_name='links') if 'links' in sheet_names else pd.DataFrame()
        stores_df = pd.read_excel(input_file, sheet_name='stores') if 'stores' in sheet_names else pd.DataFrame()
    except Exception as e:
        print(f"시트 로드 중 오류 발생: {e}")
        return
    
    # 기본 통계
    print("\n기본 통계:")
    print(f"- 버스 수: {len(buses_df)}")
    print(f"- 라인 수: {len(lines_df)}")
    print(f"- 로드 수: {len(loads_df)}")
    print(f"- 발전기 수: {len(generators_df)}")
    print(f"- 링크 수: {len(links_df) if not links_df.empty else 0}")
    print(f"- 저장소 수: {len(stores_df) if not stores_df.empty else 0}")
    
    # 버스 연결 확인
    print("\n버스 연결 상태 확인:")
    
    bus_names = set(buses_df['name'])
    
    # 각 버스의 연결 상태를 추적하는 딕셔너리
    bus_connections = {bus: {'generators': [], 'loads': [], 'lines_from': [], 
                            'lines_to': [], 'links_from': [], 'links_to': [], 
                            'stores': []} for bus in bus_names}
    
    # 발전기 연결 확인
    for _, gen in generators_df.iterrows():
        bus = gen['bus']
        if bus in bus_connections:
            bus_connections[bus]['generators'].append(gen['name'])
    
    # 로드 연결 확인
    for _, load in loads_df.iterrows():
        bus = load['bus']
        if bus in bus_connections:
            bus_connections[bus]['loads'].append(load['name'])
        else:
            print(f"경고: 로드 '{load['name']}'는 존재하지 않는 버스 '{bus}'에 연결되어 있습니다.")
    
    # 라인 연결 확인
    for _, line in lines_df.iterrows():
        bus0 = line['bus0']
        bus1 = line['bus1']
        
        if bus0 in bus_connections:
            bus_connections[bus0]['lines_from'].append(line['name'])
        else:
            print(f"경고: 라인 '{line['name']}'의 bus0 '{bus0}'가 정의되지 않았습니다.")
            
        if bus1 in bus_connections:
            bus_connections[bus1]['lines_to'].append(line['name'])
        else:
            print(f"경고: 라인 '{line['name']}'의 bus1 '{bus1}'가 정의되지 않았습니다.")
    
    # 링크 연결 확인
    if not links_df.empty:
        for _, link in links_df.iterrows():
            bus0 = link['bus0']
            bus1 = link['bus1']
            
            if bus0 in bus_connections:
                bus_connections[bus0]['links_from'].append(link['name'])
            else:
                print(f"경고: 링크 '{link['name']}'의 bus0 '{bus0}'가 정의되지 않았습니다.")
                
            if bus1 in bus_connections:
                bus_connections[bus1]['links_to'].append(link['name'])
            else:
                print(f"경고: 링크 '{link['name']}'의 bus1 '{bus1}'가 정의되지 않았습니다.")
    
    # 저장소 연결 확인
    if not stores_df.empty:
        for _, store in stores_df.iterrows():
            bus = store['bus']
            if bus in bus_connections:
                bus_connections[bus]['stores'].append(store['name'])
            else:
                print(f"경고: 저장소 '{store['name']}'는 존재하지 않는 버스 '{bus}'에 연결되어 있습니다.")
    
    # 각 버스의 균형 상태 확인
    print("\n각 버스의 균형 상태:")
    problematic_buses = []
    
    for bus, connections in bus_connections.items():
        has_generator = len(connections['generators']) > 0
        has_load = len(connections['loads']) > 0
        incoming_lines = len(connections['lines_to']) > 0
        incoming_links = len(connections['links_to']) > 0
        outgoing_lines = len(connections['lines_from']) > 0
        outgoing_links = len(connections['links_from']) > 0
        has_store = len(connections['stores']) > 0
        
        is_isolated = not (has_generator or has_load or incoming_lines or incoming_links or 
                          outgoing_lines or outgoing_links or has_store)
        
        # 로드가 있지만 입력(발전기, 들어오는 라인/링크)이 없는 경우
        is_load_without_input = has_load and not (has_generator or incoming_lines or incoming_links)
        
        # 버스 균형 확인
        if is_isolated:
            print(f"[문제] 버스 '{bus}'는 고립되어 있습니다.")
            problematic_buses.append(bus)
        elif is_load_without_input:
            print(f"[문제] 버스 '{bus}'에는 로드가 있지만, 발전기 또는 들어오는 연결이 없습니다.")
            problematic_buses.append(bus)
        else:
            print(f"[정상] 버스 '{bus}'는 적절히 연결되어 있습니다.")
            
        # 연결 세부 정보 출력
        print(f"  - 발전기: {connections['generators'] if connections['generators'] else '없음'}")
        print(f"  - 로드: {connections['loads'] if connections['loads'] else '없음'}")
        print(f"  - 들어오는 라인: {connections['lines_to'] if connections['lines_to'] else '없음'}")
        print(f"  - 나가는 라인: {connections['lines_from'] if connections['lines_from'] else '없음'}")
        if not links_df.empty:
            print(f"  - 들어오는 링크: {connections['links_to'] if connections['links_to'] else '없음'}")
            print(f"  - 나가는 링크: {connections['links_from'] if connections['links_from'] else '없음'}")
        if not stores_df.empty:
            print(f"  - 저장소: {connections['stores'] if connections['stores'] else '없음'}")
    
    # 문제가 있는 버스 요약
    print("\n문제가 있는 버스 요약:")
    if problematic_buses:
        for bus in problematic_buses:
            print(f"- {bus}")
    else:
        print("모든 버스가 정상입니다.")

if __name__ == "__main__":
    check_network_balance() 