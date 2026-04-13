import pandas as pd

def _classify_technology(gen_or_link_name):
    """PyPSA_GUI.py와 동일한 분류 로직"""
    name = str(gen_or_link_name).strip().lower()
    if 'nuclear' in name:
        return '원자력'
    if 'pv' in name or 'solar' in name:
        return '태양광'
    if 'wt' in name or 'wind' in name:
        return '풍력'
    if 'hydro' in name or 'water' in name:
        return '수력'
    # 석탄발전소 인식 - 'coal' 키워드 또는 발전소 이름으로 판별
    coal_plants = ['당진', '영흥', '하동', '삼척', '태안', '보령', '삼천포', '호남', '동해', 
                   '강릉안인', '북평', '신서천', '고성', '여수', '삼척그린파워']
    if 'coal' in name or any(plant in name for plant in coal_plants):
        return '석탄'
    if 'chp' in name:
        return 'CHP'
    if 'lng' in name or 'gas' in name:
        return 'LNG'
    if 'slack' in name or 'fallback' in name:
        # 슬랙/폴백 발전기는 연결 버스 타입에 따라 분류
        if '_h_' in name or 'heat' in name:
            return '열'
        elif '_h2_' in name or 'hydrogen' in name:
            return '수소'
        else:
            return 'LNG'  # 전력 슬랙은 LNG 기반
    if 'oil' in name or 'diesel' in name:
        return '석유'
    if 'biomass' in name or 'bio' in name:
        return '바이오'
    if 'geothermal' in name:
        return '지열'
    if 'heatpump' in name or 'heat_pump' in name or name.startswith('hp') or ' hp ' in name:
        return '히트펌프'
    if 'electrolyser' in name or 'electrolyzer' in name or 'electrolysis' in name:
        return '전해조'
    if 'h2' in name or 'hydrogen' in name:
        return '수소'
    if 'heat' in name:
        return '열'
    return '기타'

print("=" * 80)
print("'기타'로 분류되는 발전기 확인")
print("=" * 80)

df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

# 모든 발전기에 대해 분류 수행
df_gens['tech_classification'] = df_gens['name'].apply(_classify_technology)

# 기타로 분류된 발전기 필터링
other_gens = df_gens[df_gens['tech_classification'] == '기타']

print(f"\n'기타'로 분류된 발전기 수: {len(other_gens)}")
print(f"총 용량: {other_gens['p_nom'].sum():.1f} MW")

if len(other_gens) > 0:
    print(f"\n'기타' 발전기 목록:")
    print("-" * 80)
    for idx, row in other_gens.iterrows():
        print(f"  {row['name']:40s} | carrier: {row['carrier']:12s} | 용량: {row['p_nom']:8.1f} MW")
    
    # Carrier 분포
    print(f"\n'기타' 발전기의 carrier 분포:")
    carrier_counts = other_gens['carrier'].value_counts()
    for carrier, count in carrier_counts.items():
        total_cap = other_gens[other_gens['carrier'] == carrier]['p_nom'].sum()
        print(f"  {carrier:15s}: {count:3d}개, {total_cap:10.1f} MW")
    
    # carrier='coal'인 발전기 확인
    coal_others = other_gens[other_gens['carrier'] == 'coal']
    if len(coal_others) > 0:
        print(f"\n[중요!] carrier='coal'이지만 '기타'로 분류된 발전기: {len(coal_others)}개")
        print("-" * 80)
        for idx, row in coal_others.iterrows():
            print(f"  {row['name']:40s} | 용량: {row['p_nom']:8.1f} MW")
        
        print(f"\n[원인 분석]")
        print(f"  이 발전기들의 이름에는 'coal' 키워드나 발전소명이 없어서")
        print(f"  _classify_technology 함수가 '석탄'으로 인식하지 못했습니다.")
else:
    print("\n'기타'로 분류된 발전기가 없습니다.")

# 전체 분류 통계
print("\n" + "=" * 80)
print("전체 발전기 분류 통계")
print("=" * 80)

tech_counts = df_gens['tech_classification'].value_counts()
print(f"\n분류별 발전기 수:")
for tech, count in tech_counts.items():
    total_cap = df_gens[df_gens['tech_classification'] == tech]['p_nom'].sum()
    print(f"  {tech:15s}: {count:3d}개, {total_cap:10.1f} MW")

# 석탄으로 분류된 발전기 확인
coal_classified = df_gens[df_gens['tech_classification'] == '석탄']
print(f"\n'석탄'으로 분류된 발전기 샘플 (상위 5개):")
for idx, row in coal_classified.head(5).iterrows():
    print(f"  {row['name']:40s} | carrier: {row['carrier']:12s}")
