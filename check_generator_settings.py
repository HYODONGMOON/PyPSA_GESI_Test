import pandas as pd

df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

print("=" * 80)
print("1. 석탄발전기 설정 확인")
print("=" * 80)

# 석탄발전기 필터링 (세분화된 발전기들)
coal_gens = df_gens[df_gens['name'].str.contains('당진|영흥|하동|삼척|태안|보령|삼천포|호남|동해', na=False)]

print(f"\n석탄발전기 수: {len(coal_gens)}")
print(f"\nCommittable 설정:")
print(coal_gens['committable'].value_counts())

print(f"\n석탄발전기 샘플 (파라미터 확인):")
print(coal_gens[['name', 'carrier', 'committable', 'p_min_pu', 'marginal_cost', 'p_nom']].head(10))

print("\n" + "=" * 80)
print("2. LNG 발전기 설정 확인")
print("=" * 80)

lng_gens = df_gens[df_gens['name'].str.contains('LNG', na=False)]
print(f"\nLNG 발전기 수: {len(lng_gens)}")
print(f"\nCommittable 설정:")
print(lng_gens['committable'].value_counts())

print(f"\nLNG 발전기 샘플:")
print(lng_gens[['name', 'carrier', 'committable', 'p_min_pu', 'marginal_cost', 'p_nom']].head(10))

print("\n" + "=" * 80)
print("3. 모든 발전기의 Committable 설정 요약")
print("=" * 80)

print(f"\n전체 발전기 수: {len(df_gens)}")
print(f"\nCommittable 설정 분포:")
print(df_gens['committable'].value_counts())

print(f"\nCarrier별 Committable 설정:")
for carrier in ['electricity', 'gas', 'solar', 'wind', 'nuclear']:
    carrier_gens = df_gens[df_gens['carrier'] == carrier]
    if not carrier_gens.empty:
        print(f"\n  {carrier}:")
        print(f"    총 {len(carrier_gens)}개")
        if 'committable' in carrier_gens.columns:
            print(f"    Committable=True: {(carrier_gens['committable'] == True).sum()}개")
            print(f"    Committable=False: {(carrier_gens['committable'] == False).sum()}개")
