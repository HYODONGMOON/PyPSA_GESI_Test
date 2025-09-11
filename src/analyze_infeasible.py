import pandas as pd
import numpy as np

def analyze_infeasible_problem():
    """최적화 infeasible 문제 분석"""
    print("=== 최적화 Infeasible 문제 분석 ===\n")
    
    try:
        # 데이터 로드
        input_data = {}
        xls = pd.ExcelFile('integrated_input_data.xlsx')
        
        for sheet_name in xls.sheet_names:
            input_data[sheet_name] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet_name)
        
        buses_df = input_data['buses']
        generators_df = input_data['generators']
        loads_df = input_data['loads']
        links_df = input_data['links']
        lines_df = input_data['lines']
        stores_df = input_data['stores']
        
        print("1. 기본 데이터 확인:")
        print(f"   - 버스 수: {len(buses_df)}")
        print(f"   - 발전기 수: {len(generators_df)}")
        print(f"   - 부하 수: {len(loads_df)}")
        print(f"   - 링크 수: {len(links_df)}")
        print(f"   - 선로 수: {len(lines_df)}")
        print(f"   - 저장장치 수: {len(stores_df)}")
        
        # 2. 선로 연결 문제 분석
        print("\n2. 선로 연결 문제 분석:")
        bus_names = set(buses_df['name'])
        
        missing_buses_in_lines = []
        for _, line in lines_df.iterrows():
            bus0 = line['bus0']
            bus1 = line['bus1']
            
            if bus0 not in bus_names:
                missing_buses_in_lines.append(f"선로 '{line['name']}'의 bus0 '{bus0}'")
            if bus1 not in bus_names:
                missing_buses_in_lines.append(f"선로 '{line['name']}'의 bus1 '{bus1}'")
        
        if missing_buses_in_lines:
            print(f"   [문제] 존재하지 않는 버스에 연결된 선로: {len(missing_buses_in_lines)}개")
            for missing in missing_buses_in_lines[:10]:  # 처음 10개만 출력
                print(f"   - {missing}")
            if len(missing_buses_in_lines) > 10:
                print(f"   ... 및 {len(missing_buses_in_lines) - 10}개 더")
        else:
            print("   [정상] 모든 선로가 유효한 버스에 연결됨")
        
        # 3. 전력 균형 분석
        print("\n3. 전력 균형 분석:")
        
        # 총 발전 용량
        total_generation_capacity = generators_df['p_nom'].sum()
        print(f"   - 총 발전 용량: {total_generation_capacity:,.0f} MW")
        
        # 총 부하
        total_load = loads_df['p_set'].sum()
        print(f"   - 총 부하: {total_load:,.0f} MW")
        
        # 발전-부하 균형
        balance_ratio = total_generation_capacity / total_load if total_load > 0 else 0
        print(f"   - 발전/부하 비율: {balance_ratio:.2f}")
        
        if balance_ratio < 1.0:
            print("   [문제] 발전 용량이 부하보다 부족합니다!")
        elif balance_ratio > 3.0:
            print("   [경고] 발전 용량이 부하보다 과도하게 많습니다.")
        else:
            print("   [정상] 발전-부하 균형이 적절합니다.")
        
        # 4. 고립된 버스 분석
        print("\n4. 고립된 버스 분석:")
        
        # 각 버스의 연결 정보 구성
        bus_connections = {bus: {
            'generators': [],
            'loads': [],
            'incoming_lines': [],
            'outgoing_lines': [],
            'incoming_links': [],
            'outgoing_links': [],
            'stores': []
        } for bus in bus_names}
        
        # 발전기 연결
        for _, gen in generators_df.iterrows():
            bus = gen['bus']
            if bus in bus_connections:
                bus_connections[bus]['generators'].append(gen['name'])
        
        # 부하 연결
        for _, load in loads_df.iterrows():
            bus = load['bus']
            if bus in bus_connections:
                bus_connections[bus]['loads'].append(load['name'])
        
        # 선로 연결 (유효한 버스만)
        for _, line in lines_df.iterrows():
            bus0 = line['bus0']
            bus1 = line['bus1']
            
            if bus0 in bus_connections:
                bus_connections[bus0]['outgoing_lines'].append(line['name'])
            if bus1 in bus_connections:
                bus_connections[bus1]['incoming_lines'].append(line['name'])
        
        # 링크 연결
        for _, link in links_df.iterrows():
            bus0 = link['bus0']
            bus1 = link['bus1']
            
            if bus0 in bus_connections:
                bus_connections[bus0]['outgoing_links'].append(link['name'])
            if bus1 in bus_connections:
                bus_connections[bus1]['incoming_links'].append(link['name'])
        
        # 저장장치 연결
        for _, store in stores_df.iterrows():
            bus = store['bus']
            if bus in bus_connections:
                bus_connections[bus]['stores'].append(store['name'])
        
        # 문제가 있는 버스 찾기
        problematic_buses = []
        
        for bus, connections in bus_connections.items():
            has_load = len(connections['loads']) > 0
            has_generator = len(connections['generators']) > 0
            has_incoming = len(connections['incoming_lines']) > 0 or len(connections['incoming_links']) > 0
            has_outgoing = len(connections['outgoing_lines']) > 0 or len(connections['outgoing_links']) > 0
            has_store = len(connections['stores']) > 0
            
            # 부하가 있지만 공급원이 없는 버스
            if has_load and not (has_generator or has_incoming):
                problematic_buses.append({
                    'bus': bus,
                    'problem': '부하가 있지만 발전기나 들어오는 연결이 없음',
                    'loads': connections['loads'],
                    'generators': connections['generators'],
                    'incoming': connections['incoming_lines'] + connections['incoming_links']
                })
            
            # 완전히 고립된 버스
            elif not (has_load or has_generator or has_incoming or has_outgoing or has_store):
                problematic_buses.append({
                    'bus': bus,
                    'problem': '완전히 고립된 버스',
                    'loads': [],
                    'generators': [],
                    'incoming': []
                })
        
        if problematic_buses:
            print(f"   [문제] 문제가 있는 버스: {len(problematic_buses)}개")
            for i, bus_info in enumerate(problematic_buses[:5]):  # 처음 5개만 출력
                print(f"   {i+1}. 버스 '{bus_info['bus']}': {bus_info['problem']}")
                if bus_info['loads']:
                    print(f"      부하: {bus_info['loads']}")
                if bus_info['generators']:
                    print(f"      발전기: {bus_info['generators']}")
                if bus_info['incoming']:
                    print(f"      들어오는 연결: {bus_info['incoming']}")
            if len(problematic_buses) > 5:
                print(f"   ... 및 {len(problematic_buses) - 5}개 더")
        else:
            print("   [정상] 모든 버스가 적절히 연결됨")
        
        # 5. 네트워크 연결성 분석
        print("\n5. 네트워크 연결성 분석:")
        
        # 실제 연결된 선로/링크 수
        valid_lines = 0
        for _, line in lines_df.iterrows():
            if line['bus0'] in bus_names and line['bus1'] in bus_names:
                valid_lines += 1
        
        valid_links = len(links_df)  # 링크는 모두 유효한 것으로 가정
        
        print(f"   - 유효한 선로 수: {valid_lines}/{len(lines_df)}")
        print(f"   - 유효한 링크 수: {valid_links}")
        
        if valid_lines == 0 and valid_links == 0:
            print("   [심각한 문제] 버스 간 연결이 전혀 없습니다!")
        elif valid_lines < len(buses_df) - 1:
            print("   [문제] 네트워크가 완전히 연결되지 않을 수 있습니다.")
        else:
            print("   [정상] 네트워크 연결성이 적절합니다.")
        
        # 6. 해결 방안 제시
        print("\n6. 해결 방안:")
        
        if missing_buses_in_lines:
            print("   a) 선로 연결 문제 해결:")
            print("      - lines 시트의 버스 이름을 buses 시트와 일치시키기")
            print("      - 존재하지 않는 버스를 참조하는 선로 제거 또는 수정")
        
        if problematic_buses:
            print("   b) 고립된 버스 문제 해결:")
            print("      - 부하만 있는 버스에 발전기 추가 또는 다른 버스와 연결")
            print("      - 완전히 고립된 버스 제거 또는 연결 추가")
        
        if valid_lines == 0:
            print("   c) 네트워크 연결성 문제 해결:")
            print("      - 주요 버스들 간의 선로 연결 추가")
            print("      - 지역 간 송전선로 정의")
        
        print("\n분석 완료!")
        
    except Exception as e:
        print(f"분석 중 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    analyze_infeasible_problem() 