import pandas as pd

print("=" * 80)
print("송전선 용량 및 발전기 분포 확인")
print("=" * 80)

# 발전기 지역별 분포
df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

print("\n[발전기 지역별 용량 분포]")
coal_gens = df_gens[df_gens['name'].str.contains('당진|영흥|하동|삼척|태안|보령|삼천포|호남|동해', na=False)]
lng_gens = df_gens[df_gens['name'].str.contains('LNG', na=False)]

for region in ['CND', 'GND', 'ICN', 'GWD', 'BSN', 'GBD']:
    region_coal = coal_gens[coal_gens['bus'].str.startswith(region)]
    region_lng = lng_gens[lng_gens['bus'].str.startswith(region)]
    
    coal_cap = region_coal['p_nom'].sum() if len(region_coal) > 0 else 0
    lng_cap = region_lng['p_nom'].sum() if len(region_lng) > 0 else 0
    
    print(f"\n{region}:")
    print(f"  석탄: {len(region_coal)}개, {coal_cap:.1f} MW")
    print(f"  LNG: {len(region_lng)}개, {lng_cap:.1f} MW")
    print(f"  합계: {coal_cap + lng_cap:.1f} MW")

# 송전선 확인
try:
    df_lines = pd.read_excel('integrated_input_data.xlsx', sheet_name='lines')
    print("\n" + "=" * 80)
    print("[송전선 용량]")
    print("=" * 80)
    
    print(f"\n총 송전선 수: {len(df_lines)}")
    print(f"\n주요 송전선:")
    print(f"{'From':10s} -> {'To':10s} | {'용량(MW)':12s} | {'길이(km)':10s}")
    print("-" * 50)
    
    for idx, row in df_lines.iterrows():
        bus0 = row['bus0']
        bus1 = row['bus1']
        s_nom = row.get('s_nom', row.get('s_nom_min', 0))
        length = row.get('length', 0)
        print(f"{bus0:10s} -> {bus1:10s} | {s_nom:12.1f} | {length:10.1f}")
    
    # CND와 연결된 송전선 확인
    print("\n" + "=" * 80)
    print("[CND (충청남도) 관련 송전선 - 중요!]")
    print("=" * 80)
    
    cnd_lines = df_lines[(df_lines['bus0'].str.contains('CND', na=False)) | 
                         (df_lines['bus1'].str.contains('CND', na=False))]
    
    if len(cnd_lines) > 0:
        print(f"\nCND 관련 송전선: {len(cnd_lines)}개")
        print(f"총 송전 용량: {cnd_lines['s_nom'].sum():.1f} MW")
        print(f"\n상세:")
        for idx, row in cnd_lines.iterrows():
            print(f"  {row['bus0']} <-> {row['bus1']}: {row['s_nom']:.1f} MW")
        
        print(f"\n[분석]")
        print(f"  CND 석탄발전 용량: {coal_gens[coal_gens['bus'].str.startswith('CND')]['p_nom'].sum():.1f} MW")
        print(f"  CND 관련 송전선 합계: {cnd_lines['s_nom'].sum():.1f} MW")
        
        cnd_coal = coal_gens[coal_gens['bus'].str.startswith('CND')]['p_nom'].sum()
        cnd_transmission = cnd_lines['s_nom'].sum()
        
        if cnd_coal > cnd_transmission:
            print(f"\n  [경고!] CND의 석탄발전 용량({cnd_coal:.1f} MW)이")
            print(f"          송전선 용량({cnd_transmission:.1f} MW)을 초과합니다!")
            print(f"          이것이 infeasible의 원인일 수 있습니다.")
    else:
        print("\n  CND 관련 송전선이 없습니다!")
        print("  [심각한 문제] CND는 독립된 섬(Island)일 수 있습니다.")
    
except Exception as e:
    print(f"\n송전선 데이터 읽기 실패: {e}")
