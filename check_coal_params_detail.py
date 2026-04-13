import pandas as pd

df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

# 석탄발전기 필터링
coal_gens = df_gens[df_gens['name'].str.contains('당진|영흥|하동|삼척|태안|보령|삼천포|호남|동해', na=False)]

print("=" * 80)
print("석탄발전기 파라미터 상세 확인")
print("=" * 80)

print(f"\n석탄발전기 수: {len(coal_gens)}")
print(f"총 용량: {coal_gens['p_nom'].sum():.1f} MW")

# 모든 열 확인
print(f"\n컬럼 목록:")
print(df_gens.columns.tolist())

# 주요 파라미터 확인
important_cols = ['name', 'bus', 'carrier', 'p_nom', 'efficiency', 'marginal_cost', 
                  'committable', 'p_min_pu', 'p_max_pu', 'ramp_limit_up', 'ramp_limit_down',
                  'capital_cost', 'p_nom_extendable', 'p_nom_min', 'p_nom_max']

existing_cols = [col for col in important_cols if col in df_gens.columns]

print(f"\n석탄발전기 파라미터 샘플 (상위 5개):")
print(coal_gens[existing_cols].head(5))

# 문제될 수 있는 파라미터 체크
print("\n" + "=" * 80)
print("잠재적 문제 체크")
print("=" * 80)

# 1. p_min_pu 확인
if 'p_min_pu' in coal_gens.columns:
    print("\n1. p_min_pu 설정:")
    print(f"   값 분포: {coal_gens['p_min_pu'].value_counts().to_dict()}")
    non_zero_p_min = coal_gens[coal_gens['p_min_pu'] > 0]
    if len(non_zero_p_min) > 0:
        print(f"   [주의] p_min_pu > 0인 발전기: {len(non_zero_p_min)}개")
        print(f"   최소 출력 제약이 있는 발전기들: {non_zero_p_min['p_min_pu'].unique()}")

# 2. ramp_limit 확인
if 'ramp_limit_up' in coal_gens.columns:
    print("\n2. ramp_limit_up 설정:")
    print(f"   값 분포: {coal_gens['ramp_limit_up'].value_counts().to_dict()}")

# 3. p_nom_extendable 확인
if 'p_nom_extendable' in coal_gens.columns:
    print("\n3. p_nom_extendable 설정:")
    print(f"   값 분포: {coal_gens['p_nom_extendable'].value_counts().to_dict()}")

# 4. 버스 연결 확인
print("\n4. 버스 연결:")
bus_counts = coal_gens['bus'].value_counts()
for bus, count in bus_counts.items():
    print(f"   {bus}: {count}개")

# 5. LNG와 비교
lng_gens = df_gens[df_gens['name'].str.contains('LNG', na=False)]
print("\n" + "=" * 80)
print("LNG 발전기와 비교")
print("=" * 80)

print(f"\nLNG 발전기 샘플 (상위 3개):")
print(lng_gens[existing_cols].head(3))
