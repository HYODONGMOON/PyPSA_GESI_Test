import pandas as pd

def check_line_structure():
    print("=== 송전선로 구조 분석 ===")
    
    # 송전선로 데이터 로드
    lines_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='lines')
    
    print(f"전체 송전선로: {len(lines_df)}개")
    
    # ICN-GGD 연결 선로
    print("\n=== ICN-GGD 연결 선로 ===")
    icn_ggd = lines_df[
        ((lines_df['bus0'].str.contains('ICN', na=False)) & (lines_df['bus1'].str.contains('GGD', na=False))) |
        ((lines_df['bus0'].str.contains('GGD', na=False)) & (lines_df['bus1'].str.contains('ICN', na=False)))
    ]
    
    for _, line in icn_ggd.iterrows():
        print(f"{line['name']}: {line['bus0']} ↔ {line['bus1']} ({line['s_nom']:.1f} MW)")
    
    print(f"ICN-GGD 총 용량: {icn_ggd['s_nom'].sum():.1f} MW")
    
    # GGD-CND 연결 선로  
    print("\n=== GGD-CND 연결 선로 ===")
    ggd_cnd = lines_df[
        ((lines_df['bus0'].str.contains('GGD', na=False)) & (lines_df['bus1'].str.contains('CND', na=False))) |
        ((lines_df['bus0'].str.contains('CND', na=False)) & (lines_df['bus1'].str.contains('GGD', na=False)))
    ]
    
    for _, line in ggd_cnd.iterrows():
        print(f"{line['name']}: {line['bus0']} ↔ {line['bus1']} ({line['s_nom']:.1f} MW)")
    
    print(f"GGD-CND 총 용량: {ggd_cnd['s_nom'].sum():.1f} MW")
    
    # GND-DGU 연결 선로
    print("\n=== GND-DGU 연결 선로 ===")
    gnd_dgu = lines_df[
        ((lines_df['bus0'].str.contains('GND', na=False)) & (lines_df['bus1'].str.contains('DGU', na=False))) |
        ((lines_df['bus0'].str.contains('DGU', na=False)) & (lines_df['bus1'].str.contains('GND', na=False)))
    ]
    
    for _, line in gnd_dgu.iterrows():
        print(f"{line['name']}: {line['bus0']} ↔ {line['bus1']} ({line['s_nom']:.1f} MW)")
    
    print(f"GND-DGU 총 용량: {gnd_dgu['s_nom'].sum():.1f} MW")
    
    # 지역간 연결 요약
    print("\n=== 지역간 연결 요약 ===")
    regional_connections = {}
    
    for _, line in lines_df.iterrows():
        bus0_region = line['bus0'].split('_')[0] if '_' in line['bus0'] else line['bus0'][:3]
        bus1_region = line['bus1'].split('_')[0] if '_' in line['bus1'] else line['bus1'][:3]
        
        if bus0_region != bus1_region:
            connection = f"{bus0_region}-{bus1_region}"
            reverse_connection = f"{bus1_region}-{bus0_region}"
            
            # 정규화된 연결명 (알파벳 순)
            normalized = connection if bus0_region < bus1_region else reverse_connection
            
            if normalized not in regional_connections:
                regional_connections[normalized] = {
                    'lines': [],
                    'total_capacity': 0
                }
            
            regional_connections[normalized]['lines'].append({
                'name': line['name'],
                'capacity': line['s_nom']
            })
            regional_connections[normalized]['total_capacity'] += line['s_nom']
    
    # 용량별 정렬
    sorted_connections = sorted(regional_connections.items(), key=lambda x: x[1]['total_capacity'], reverse=True)
    
    print("주요 지역간 연결 (용량 순):")
    for connection, data in sorted_connections[:15]:
        print(f"\n{connection}: {data['total_capacity']:.1f} MW ({len(data['lines'])}개 선로)")
        for line_info in data['lines']:
            print(f"  - {line_info['name']}: {line_info['capacity']:.1f} MW")

if __name__ == "__main__":
    check_line_structure() 