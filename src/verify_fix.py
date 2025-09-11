import pandas as pd

def verify_fix():
    """문제가 해결되었는지 확인하는 스크립트"""
    
    # 파일 로드
    input_file = 'integrated_input_data.xlsx'
    print(f"파일 분석: {input_file}")
    
    # 각 시트 로드
    buses_df = pd.read_excel(input_file, sheet_name='buses')
    links_df = pd.read_excel(input_file, sheet_name='links')
    loads_df = pd.read_excel(input_file, sheet_name='loads')
    
    print(f"버스 수: {len(buses_df)}")
    print(f"링크 수: {len(links_df)}")
    print(f"로드 수: {len(loads_df)}")
    
    # 문제가 있던 버스들 확인
    problematic_buses = ['SEL_Hydrogen', 'JBD_Hydrogen']
    print("\n이전에 문제가 있던 버스들 확인:")
    
    for bus_name in problematic_buses:
        # 버스 존재 확인
        bus_exists = bus_name in buses_df['name'].values
        
        if bus_exists:
            # 해당 버스로 오는 링크 확인
            incoming_links = links_df[links_df['bus1'] == bus_name]
            
            # 해당 버스의 로드 확인
            bus_loads = loads_df[loads_df['bus'] == bus_name]
            
            print(f"\n버스: {bus_name}")
            print(f"  존재 여부: {bus_exists}")
            print(f"  들어오는 링크 수: {len(incoming_links)}")
            
            if not incoming_links.empty:
                print("  들어오는 링크:")
                for _, link in incoming_links.iterrows():
                    print(f"    - {link['name']} (from {link['bus0']})")
            
            print(f"  로드 수: {len(bus_loads)}")
            if not bus_loads.empty:
                print("  로드:")
                for _, load in bus_loads.iterrows():
                    print(f"    - {load['name']}, p_set: {load['p_set']}")
            
            # 문제 해결 확인
            if len(incoming_links) > 0:
                print("  [해결] 버스에 들어오는 링크가 있습니다.")
            else:
                print("  [문제 지속] 버스에 들어오는 링크가 없습니다!")
        else:
            print(f"\n버스: {bus_name}")
            print("  존재 하지 않음 (삭제됨)")
    
    # 새로 추가된 링크 확인
    new_links = ['SEL_Electrolyser', 'JBD_Electrolyser']
    print("\n새로 추가된 링크 확인:")
    
    for link_name in new_links:
        link_exists = link_name in links_df['name'].values
        
        if link_exists:
            link_info = links_df[links_df['name'] == link_name].iloc[0]
            print(f"링크: {link_name}")
            print(f"  존재 여부: {link_exists}")
            print(f"  bus0: {link_info['bus0']}")
            print(f"  bus1: {link_info['bus1']}")
            print(f"  carrier: {link_info['carrier']}")
            print(f"  p_nom: {link_info['p_nom']}")
        else:
            print(f"링크: {link_name}")
            print("  존재 하지 않음")
    
    # 종합 결과
    print("\n종합 결과:")
    all_fixed = True
    
    for bus_name in problematic_buses:
        if bus_name in buses_df['name'].values:
            incoming_links = links_df[links_df['bus1'] == bus_name]
            if len(incoming_links) == 0:
                all_fixed = False
                print(f"  - {bus_name} 버스에 들어오는 링크가 없습니다!")
    
    if all_fixed:
        print("  모든 문제가 해결되었습니다! 네트워크 최적화를 시도해 보세요.")
    else:
        print("  일부 문제가 해결되지 않았습니다. 추가 수정이 필요합니다.")

if __name__ == "__main__":
    verify_fix() 