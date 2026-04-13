import pandas as pd

print('='*70)
print('시트 및 Links 확인')
print('='*70)

# 1. 시트 목록
sheets = pd.ExcelFile('integrated_input_data.xlsx').sheet_names
print('\n1. 시트 목록:')
for sheet in sheets:
    print(f'   - {sheet}')

# 2. Links 시트
print('\n2. Links 시트:')
try:
    links_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='links')
    print(f'   - 총 {len(links_df)}개')
    
    if len(links_df) > 0:
        print(f'\n   처음 10개:')
        print(links_df[['name', 'bus0', 'bus1', 'carrier']].head(10).to_string(index=False))
        
        print(f'\n   Carrier 종류:')
        print(links_df['carrier'].value_counts().head(20).to_string())
        
        # CHP 관련 links
        chp_links = links_df[links_df['carrier'].str.contains('CHP|chp|열병합', case=False, na=False)]
        print(f'\n   CHP 관련 links: {len(chp_links)}개')
        
        # Heat 관련 links (bus0 or bus1에 _H 포함)
        heat_links = links_df[links_df['bus0'].str.contains('_H', na=False) | links_df['bus1'].str.contains('_H', na=False)]
        print(f'   Heat 관련 links: {len(heat_links)}개')
except Exception as e:
    print(f'   오류: {e}')

# 3. Generators 시트에서 carrier 재확인
print('\n3. Generators 시트 상세:')
gens_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
print(f'   - 총 {len(gens_df)}개')
print(f'\n   Carrier 종류:')
print(gens_df['carrier'].value_counts().to_string())

print(f'\n   처음 20개:')
print(gens_df[['name', 'bus', 'carrier', 'p_nom']].head(20).to_string(index=False))

print('\n' + '='*70)
