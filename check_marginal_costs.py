import pandas as pd

df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

print("=" * 80)
print("발전기별 marginal_cost 확인")
print("=" * 80)

# 석탄발전기
coal_gens = df_gens[df_gens['name'].str.contains('당진|영흥|하동|삼척|태안|보령|삼천포|호남|동해', na=False)]
print(f"\n석탄발전기:")
print(f"  수: {len(coal_gens)}")
print(f"  marginal_cost 범위: {coal_gens['marginal_cost'].min():.2f} ~ {coal_gens['marginal_cost'].max():.2f}")
print(f"  평균 marginal_cost: {coal_gens['marginal_cost'].mean():.2f}")
print(f"  샘플:")
for idx, row in coal_gens.head(5).iterrows():
    print(f"    {row['name']}: {row['marginal_cost']:.2f} KRW/MWh, 효율: {row['efficiency']:.3f}")

# LNG발전기
lng_gens = df_gens[df_gens['name'].str.contains('LNG', na=False)]
print(f"\nLNG발전기:")
print(f"  수: {len(lng_gens)}")
print(f"  marginal_cost 범위: {lng_gens['marginal_cost'].min():.2f} ~ {lng_gens['marginal_cost'].max():.2f}")
print(f"  평균 marginal_cost: {lng_gens['marginal_cost'].mean():.2f}")
print(f"  샘플:")
for idx, row in lng_gens.head(3).iterrows():
    print(f"    {row['name']}: {row['marginal_cost']:.2f} KRW/MWh, 효율: {row['efficiency']:.3f}")

# 다른 발전기들
other_gens = df_gens[~df_gens['name'].str.contains('당진|영흥|하동|삼척|태안|보령|삼천포|호남|동해|LNG', na=False)]
print(f"\n기타 발전기:")
print(f"  수: {len(other_gens)}")
if len(other_gens) > 0:
    print(f"  marginal_cost 범위: {other_gens['marginal_cost'].min():.2f} ~ {other_gens['marginal_cost'].max():.2f}")
    print(f"  carrier 분포:")
    for carrier, group in other_gens.groupby('carrier'):
        print(f"    {carrier}: {len(group)}개, marginal_cost 평균: {group['marginal_cost'].mean():.2f}")

print("\n" + "=" * 80)
print("Merit Order 분석")
print("=" * 80)

# 발전기를 marginal_cost 기준으로 정렬
df_sorted = df_gens.sort_values('marginal_cost')
print(f"\nMarginal cost 기준 발전 순서:")
for idx, row in df_sorted.head(15).iterrows():
    carrier = row['carrier']
    name = row['name']
    mc = row['marginal_cost']
    eff = row['efficiency']
    capacity = row['p_nom']
    print(f"  {idx+1}. {name[:30]:30s} | carrier: {carrier:10s} | MC: {mc:8.2f} | 효율: {eff:.3f} | 용량: {capacity:8.1f} MW")
