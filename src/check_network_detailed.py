import pandas as pd
import numpy as np
import os

# 입력 파일 경로
input_file = 'integrated_input_data.xlsx'

def check_network_balance():
    print("네트워크 균형 상태 상세 확인 중...")
    
    # 입력 데이터 로드
    if not os.path.exists(input_file):
        print(f"오류: {input_file} 파일이 존재하지 않습니다.")
        return
    
    # 각 시트 로드
    buses = pd.read_excel(input_file, sheet_name='buses')
    generators = pd.read_excel(input_file, sheet_name='generators')
    loads = pd.read_excel(input_file, sheet_name='loads')
    lines = pd.read_excel(input_file, sheet_name='lines') if 'lines' in pd.ExcelFile(input_file).sheet_names else pd.DataFrame()
    stores = pd.read_excel(input_file, sheet_name='stores') if 'stores' in pd.ExcelFile(input_file).sheet_names else pd.DataFrame()
    links = pd.read_excel(input_file, sheet_name='links') if 'links' in pd.ExcelFile(input_file).sheet_names else pd.DataFrame()
    
    print(f"총 버스 수: {len(buses)}")
    print(f"총 발전기 수: {len(generators)}")
    print(f"총 부하 수: {len(loads)}")
    print(f"총 라인 수: {len(lines) if not lines.empty else 0}")
    print(f"총 저장장치 수: {len(stores) if not stores.empty else 0}")
    print(f"총 링크 수: {len(links) if not links.empty else 0}")
    
    # 각 버스의 연결 상태 확인
    bus_connections = {}
    for idx, bus_row in buses.iterrows():
        bus_name = bus_row['name']
        bus_connections[bus_name] = {
            'bus_idx': idx,
            'carrier': bus_row['carrier'],
            'generators': generators[generators['bus'] == bus_name]['name'].tolist(),
            'loads': loads[loads['bus'] == bus_name]['name'].tolist(),
            'stores': stores[stores['bus'] == bus_name]['name'].tolist() if not stores.empty else [],
            'outgoing_lines': lines[lines['bus0'] == bus_name]['name'].tolist() if not lines.empty else [],
            'incoming_lines': lines[lines['bus1'] == bus_name]['name'].tolist() if not lines.empty else [],
            'outgoing_links': links[links['bus0'] == bus_name]['name'].tolist() if not links.empty else [],
            'incoming_links': links[links['bus1'] == bus_name]['name'].tolist() if not links.empty else []
        }
    
    # 모든 버스 연결 상태 자세히 출력
    print("\n모든 버스 연결 상태:")
    for bus_name, connections in bus_connections.items():
        print(f"\n버스: {bus_name} (캐리어: {connections['carrier']})")
        has_generator = len(connections['generators']) > 0
        has_load = len(connections['loads']) > 0
        has_store = len(connections['stores']) > 0
        has_incoming = len(connections['incoming_lines']) > 0 or len(connections['incoming_links']) > 0
        has_outgoing = len(connections['outgoing_lines']) > 0 or len(connections['outgoing_links']) > 0
        
        # 연결 정보 출력
        print(f"  발전기: {connections['generators'] if has_generator else '없음'}")
        print(f"  부하: {connections['loads'] if has_load else '없음'}")
        print(f"  저장장치: {connections['stores'] if has_store else '없음'}")
        print(f"  나가는 라인: {connections['outgoing_lines'] if has_outgoing else '없음'}")
        print(f"  들어오는 라인: {connections['incoming_lines'] if has_incoming else '없음'}")
        print(f"  나가는 링크: {connections['outgoing_links'] if connections['outgoing_links'] else '없음'}")
        print(f"  들어오는 링크: {connections['incoming_links'] if connections['incoming_links'] else '없음'}")
        
        # 균형 상태 체크
        if has_load and not (has_generator or has_incoming or has_store):
            print(f"  [문제] 부하가 있지만 발전기나 들어오는 연결이 없습니다.")
        elif not (has_load or has_generator or has_incoming or has_outgoing or has_store):
            print(f"  [문제] 고립된 버스입니다. 연결된 요소가 없습니다.")
        else:
            print("  [정상] 버스 연결이 정상입니다.")
    
    # 문제가 있는 버스 검색
    problem_buses = []
    for bus_name, connections in bus_connections.items():
        # 부하가 있지만 발전이나 연결이 없는 버스
        has_generator = len(connections['generators']) > 0
        has_load = len(connections['loads']) > 0
        has_store = len(connections['stores']) > 0
        has_incoming = len(connections['incoming_lines']) > 0 or len(connections['incoming_links']) > 0
        has_outgoing = len(connections['outgoing_lines']) > 0 or len(connections['outgoing_links']) > 0
        
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
    
    # 문제 요약
    print("\n\n문제가 있는 버스 요약:")
    if problem_buses:
        for problem in problem_buses:
            print(f"\n버스: {problem['bus']}")
            print(f"문제: {problem['reason']}")
    else:
        print("모든 버스가 균형 상태입니다.")
    
    # 부하 정보 확인
    print("\n\n부하 정보:")
    for idx, load_row in loads.iterrows():
        bus_name = load_row['bus']
        print(f"{load_row['name']}: {load_row['p_set']} (버스: {bus_name}, 캐리어: {load_row['carrier']})")
        
        # 해당 버스에 대한 연결 정보 확인
        if bus_name in bus_connections:
            connections = bus_connections[bus_name]
            has_generator = len(connections['generators']) > 0
            has_incoming = len(connections['incoming_lines']) > 0 or len(connections['incoming_links']) > 0
            has_store = len(connections['stores']) > 0
            
            if not (has_generator or has_incoming or has_store):
                print(f"  [문제] 이 부하를 만족시킬 수 있는 발전기나 연결이 없습니다.")
            else:
                print("  [정상] 이 부하는 정상적으로 만족될 수 있습니다.")
        else:
            print(f"  [심각한 문제] 이 부하가 연결된 버스({bus_name})가 존재하지 않습니다.")
    
    # 라인 정보 확인
    if not lines.empty:
        print("\n\n라인 정보:")
        for idx, line_row in lines.iterrows():
            print(f"{line_row['name']}: {line_row['bus0']} -> {line_row['bus1']} (x = {line_row['x']})")
            
            # x가 0인 라인 확인
            if line_row['x'] == 0:
                print(f"  [경고] 이 라인의 x 값이 0입니다.")
            
            # 버스 존재 여부 확인
            if line_row['bus0'] not in bus_connections:
                print(f"  [심각한 문제] 시작 버스({line_row['bus0']})가 존재하지 않습니다.")
            if line_row['bus1'] not in bus_connections:
                print(f"  [심각한 문제] 도착 버스({line_row['bus1']})가 존재하지 않습니다.")

if __name__ == "__main__":
    check_network_balance() 