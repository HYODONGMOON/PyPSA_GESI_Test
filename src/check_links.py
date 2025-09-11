import pandas as pd

def check_links():
    """링크 연결 문제 검사를 위한 스크립트"""
    
    # 파일 로드
    input_file = 'integrated_input_data_backup_20250519_134702.xlsx'
    print(f"파일 분석: {input_file}")
    
    # 각 시트 로드
    buses_df = pd.read_excel(input_file, sheet_name='buses')
    links_df = pd.read_excel(input_file, sheet_name='links')
    
    print(f"버스 수: {len(buses_df)}")
    print(f"링크 수: {len(links_df)}")
    
    # 1. 모든 링크의 bus0/bus1 존재 여부 확인
    print("\n1. 모든 링크의 버스 존재 여부 확인:")
    bus_names = set(buses_df['name'])
    
    problematic_links = []
    for idx, link in links_df.iterrows():
        bus0 = link['bus0']
        bus1 = link['bus1']
        
        if bus0 not in bus_names or bus1 not in bus_names:
            problematic_links.append({
                'name': link['name'],
                'bus0': bus0,
                'bus0_exists': bus0 in bus_names,
                'bus1': bus1,
                'bus1_exists': bus1 in bus_names
            })
    
    if problematic_links:
        print(f"  문제가 있는 링크 수: {len(problematic_links)}")
        for link in problematic_links:
            print(f"  - 링크: {link['name']}")
            if not link['bus0_exists']:
                print(f"    [문제] bus0 '{link['bus0']}'가 정의되지 않았습니다.")
            if not link['bus1_exists']:
                print(f"    [문제] bus1 '{link['bus1']}'가 정의되지 않았습니다.")
    else:
        print("  모든 링크의 버스가 제대로 정의되어 있습니다.")
    
    # 2. 링크 연결 확인 - 문제가 있는 버스 대상으로 확인
    problem_buses = [
        'JBD_Hydrogen', 'SEL_Hydrogen', 'BSN_H2', 'JJD_H2', 'GWJ_H2', 'USN_H2', 
        'SJN_H2', 'DJN_H2', 'GND_H2', 'GWD_H2', 'CND_H2', 'GGD_H2', 'GBD_H2',
        'CBD_H2', 'DGU_H2', 'ICN_H2', 'JND_H2',
        'BSN_H', 'JJD_H', 'GWJ_H', 'USN_H', 'SJN_H', 'DJN_H', 'GND_H', 'GWD_H',
        'CND_H', 'GGD_H', 'GBD_H', 'CBD_H', 'DGU_H', 'ICN_H', 'JND_H', 'SEL_H', 'JBD_H'
    ]
    
    print("\n2. 문제 버스의 링크 연결 확인:")
    
    # 각 문제 버스에 대한 링크 확인
    for bus in problem_buses:
        # 해당 버스가 bus0인 링크
        outgoing_links = links_df[links_df['bus0'] == bus]
        # 해당 버스가 bus1인 링크
        incoming_links = links_df[links_df['bus1'] == bus]
        
        print(f"\n버스: {bus}")
        print(f"  들어오는 링크 수: {len(incoming_links)}")
        if not incoming_links.empty:
            for _, link in incoming_links.iterrows():
                print(f"    - {link['name']} (from {link['bus0']})")
        else:
            print("    - 없음")
        
        print(f"  나가는 링크 수: {len(outgoing_links)}")
        if not outgoing_links.empty:
            for _, link in outgoing_links.iterrows():
                print(f"    - {link['name']} (to {link['bus1']})")
        else:
            print("    - 없음")
    
    # 3. 모든 버스에 대한 링크 연결 확인
    print("\n3. 버스별 링크 연결 요약:")
    bus_link_counts = {bus: {'incoming': 0, 'outgoing': 0} for bus in bus_names}
    
    for _, link in links_df.iterrows():
        bus0 = link['bus0']
        bus1 = link['bus1']
        
        if bus0 in bus_link_counts:
            bus_link_counts[bus0]['outgoing'] += 1
        if bus1 in bus_link_counts:
            bus_link_counts[bus1]['incoming'] += 1
    
    # 링크 연결이 있는 버스 수
    buses_with_links = sum(1 for bus, counts in bus_link_counts.items() 
                        if counts['incoming'] > 0 or counts['outgoing'] > 0)
    print(f"  링크 연결이 있는 버스 수: {buses_with_links}/{len(buses_df)}")
    
    # 입력 링크가 있는 문제 버스
    problem_buses_with_incoming_links = [
        bus for bus in problem_buses if bus in bus_link_counts and bus_link_counts[bus]['incoming'] > 0
    ]
    print(f"  입력 링크가 있는 문제 버스 수: {len(problem_buses_with_incoming_links)}/{len(problem_buses)}")
    if problem_buses_with_incoming_links:
        print("  입력 링크가 있는 문제 버스:")
        for bus in problem_buses_with_incoming_links:
            print(f"    - {bus} (들어오는 링크: {bus_link_counts[bus]['incoming']})")

if __name__ == "__main__":
    check_links() 