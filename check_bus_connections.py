import pandas as pd

print("=" * 80)
print("버스 연결 상태 확인")
print("=" * 80)

# 1. buses 확인
df_buses = pd.read_excel('integrated_input_data.xlsx', sheet_name='buses')
print(f"\n[buses 시트]")
print(f"총 버스 수: {len(df_buses)}")
print(f"버스 목록:")
for idx, row in df_buses.iterrows():
    print(f"  {row['name']}: carrier={row.get('carrier', 'N/A')}")

# 2. 석탄발전기가 연결된 버스 확인
df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
coal_gens = df_gens[df_gens['name'].str.contains('당진|영흥|하동|삼척|태안|보령|삼천포|호남|동해', na=False)]

print("\n" + "=" * 80)
print("[석탄발전기 버스 연결]")
print("=" * 80)

bus_list = df_buses['name'].tolist()

for idx, row in coal_gens.iterrows():
    gen_name = row['name']
    gen_bus = row['bus']
    
    # 버스가 존재하는지 확인
    if gen_bus in bus_list:
        status = "[OK]"
    else:
        status = "[ERROR - 버스 없음!]"
    
    print(f"{status} {gen_name[:30]:30s} -> {gen_bus}")

# 3. 버스별 석탄발전기 집계
print("\n" + "=" * 80)
print("[버스별 석탄발전기 집계]")
print("=" * 80)

for bus in coal_gens['bus'].unique():
    bus_coal = coal_gens[coal_gens['bus'] == bus]
    total_cap = bus_coal['p_nom'].sum()
    
    # 버스 존재 여부
    exists = "존재" if bus in bus_list else "존재하지 않음"
    
    print(f"\n{bus} ({exists}):")
    print(f"  발전기 수: {len(bus_coal)}")
    print(f"  총 용량: {total_cap:.1f} MW")
    print(f"  발전기 목록:")
    for idx, gen in bus_coal.iterrows():
        print(f"    - {gen['name']}: {gen['p_nom']:.1f} MW")

# 4. 이전 통합 Coal 발전기가 남아있는지 확인
print("\n" + "=" * 80)
print("[통합 Coal 발전기 확인]")
print("=" * 80)

# 'Coal'이라는 이름을 가진 발전기 찾기
coal_keyword_gens = df_gens[df_gens['name'].str.contains('Coal', case=False, na=False)]

if len(coal_keyword_gens) > 0:
    print(f"\n[경고!] 'Coal' 키워드를 포함한 발전기 발견: {len(coal_keyword_gens)}개")
    print("이전 통합 발전기가 제거되지 않았을 수 있습니다!")
    for idx, row in coal_keyword_gens.iterrows():
        print(f"  - {row['name']}: {row['p_nom']:.1f} MW, bus={row['bus']}")
else:
    print("\n[OK] 'Coal' 키워드를 포함한 발전기 없음 (정상)")

# 5. 같은 버스에 연결된 다른 발전기들 확인 (비교를 위해)
print("\n" + "=" * 80)
print("[비교: LNG 발전기 버스 연결]")
print("=" * 80)

lng_gens = df_gens[df_gens['name'].str.contains('LNG', na=False)]
for idx, row in lng_gens.iterrows():
    gen_name = row['name']
    gen_bus = row['bus']
    
    if gen_bus in bus_list:
        status = "[OK]"
    else:
        status = "[ERROR]"
    
    print(f"{status} {gen_name:15s} -> {gen_bus}")
