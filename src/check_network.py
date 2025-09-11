import pandas as pd
import numpy as np
import os

# 입력 파일 경로
input_file = 'integrated_input_data.xlsx'

def check_network_balance():
    print("네트워크 균형 상태 확인 중...")
    
    # 입력 데이터 로드
    if not os.path.exists(input_file):
        print(f"오류: {input_file} 파일이 존재하지 않습니다.")
        return
    
    # 각 시트 로드
    buses = pd.read_excel(input_file, sheet_name='buses')
    generators = pd.read_excel(input_file, sheet_name='generators')
    loads = pd.read_excel(input_file, sheet_name='loads')
    lines = pd.read_excel(input_file, sheet_name='lines')
    stores = pd.read_excel(input_file, sheet_name='stores')
    links = pd.read_excel(input_file, sheet_name='links') if 'links' in pd.ExcelFile(input_file).sheet_names else pd.DataFrame()
    
    print(f"총 버스 수: {len(buses)}")
    print(f"총 발전기 수: {len(generators)}")
    print(f"총 부하 수: {len(loads)}")
    print(f"총 라인 수: {len(lines)}")
    print(f"총 저장장치 수: {len(stores)}")
    print(f"총 링크 수: {len(links)}")
    
    # 각 버스의 연결 상태 확인
    bus_connections = {}
    for bus_name in buses['name']:
        bus_connections[bus_name] = {
            'generators': generators[generators['bus'] == bus_name]['name'].tolist(),
            'loads': loads[loads['bus'] == bus_name]['name'].tolist(),
            'stores': stores[stores['bus'] == bus_name]['name'].tolist() if not stores.empty else [],
            'outgoing_lines': lines[lines['bus0'] == bus_name]['name'].tolist(),
            'incoming_lines': lines[lines['bus1'] == bus_name]['name'].tolist(),
            'outgoing_links': links[links['bus0'] == bus_name]['name'].tolist() if not links.empty else [],
            'incoming_links': links[links['bus1'] == bus_name]['name'].tolist() if not links.empty else []
        }
    
    # 문제가 있는 버스 찾기
    problem_buses = []
    for bus_name, connections in bus_connections.items():
        # 부하가 있지만 발전이나 연결이 없는 버스
        has_load = len(connections['loads']) > 0
        has_generator = len(connections['generators']) > 0
        has_incoming = len(connections['incoming_lines']) > 0 or len(connections['incoming_links']) > 0
        has_outgoing = len(connections['outgoing_lines']) > 0 or len(connections['outgoing_links']) > 0
        has_store = len(connections['stores']) > 0
        
        if has_load and not (has_generator or has_incoming or has_store):
            problem_buses.append({
                'bus': bus_name,
                'reason': '부하가 있지만 발전기나 들어오는 연결이 없습니다.',
                'connections': connections
            })
        elif not (has_load or has_generator or has_incoming or has_outgoing or has_store):
            problem_buses.append({
                'bus': bus_name,
                'reason': '고립된 버스입니다. 연결된 요소가 없습니다.',
                'connections': connections
            })
    
    # 문제 버스 출력
    if problem_buses:
        print("\n문제가 있는 버스:")
        for problem in problem_buses:
            print(f"\n버스: {problem['bus']}")
            print(f"문제: {problem['reason']}")
            print(f"연결된 발전기: {problem['connections']['generators']}")
            print(f"연결된 부하: {problem['connections']['loads']}")
            print(f"연결된 저장장치: {problem['connections']['stores']}")
            print(f"나가는 라인: {problem['connections']['outgoing_lines']}")
            print(f"들어오는 라인: {problem['connections']['incoming_lines']}")
            print(f"나가는 링크: {problem['connections']['outgoing_links']}")
            print(f"들어오는 링크: {problem['connections']['incoming_links']}")
    else:
        print("\n모든 버스가 균형 상태입니다.")
    
    # JJD_JND 라인 확인
    jjd_jnd_line = lines[lines['name'] == 'JJD_JND']
    if not jjd_jnd_line.empty:
        print("\nJJD_JND 라인 상태:")
        print(jjd_jnd_line)
        
        # 연결된 버스 확인
        bus0 = jjd_jnd_line['bus0'].iloc[0]
        bus1 = jjd_jnd_line['bus1'].iloc[0]
        
        print(f"\nbus0 ({bus0}) 상태:")
        if bus0 in bus_connections:
            connections = bus_connections[bus0]
            print(f"연결된 발전기: {connections['generators']}")
            print(f"연결된 부하: {connections['loads']}")
        else:
            print(f"버스 {bus0}를 찾을 수 없습니다.")
            
        print(f"\nbus1 ({bus1}) 상태:")
        if bus1 in bus_connections:
            connections = bus_connections[bus1]
            print(f"연결된 발전기: {connections['generators']}")
            print(f"연결된 부하: {connections['loads']}")
        else:
            print(f"버스 {bus1}를 찾을 수 없습니다.")
    else:
        print("\nJJD_JND 라인을 찾을 수 없습니다.")

if __name__ == "__main__":
    check_network_balance() 