import pandas as pd

print("=" * 80)
print("CO2 제약 확인")
print("=" * 80)

# 1. interface.xlsx에서 CO2 설정 확인
try:
    df_co2 = pd.read_excel('interface.xlsx', sheet_name='CO2_limit')
    print("\n[interface.xlsx의 CO2_limit 시트]")
    print(df_co2)
except Exception as e:
    print(f"\n[경고] CO2_limit 시트 읽기 실패: {e}")

# 2. integrated_input_data.xlsx에서 확인
try:
    df_co2_int = pd.read_excel('integrated_input_data.xlsx', sheet_name='CO2_limit')
    print("\n[integrated_input_data.xlsx의 CO2_limit 시트]")
    print(df_co2_int)
except Exception as e:
    print(f"\n[경고] integrated_input_data.xlsx의 CO2_limit 시트 읽기 실패: {e}")

# 3. 발전기별 CO2 배출량 계산
df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

print("\n" + "=" * 80)
print("발전기별 CO2 배출 파라미터")
print("=" * 80)

# carrier별 발전기 수 확인
print("\ncarrier별 발전기 분포:")
carrier_counts = df_gens['carrier'].value_counts()
for carrier, count in carrier_counts.items():
    print(f"  {carrier}: {count}개")

# 석탄과 LNG 발전기의 효율 및 배출계수 확인
coal_gens = df_gens[df_gens['name'].str.contains('당진|영흥|하동|삼척|태안|보령|삼천포|호남|동해', na=False)]
lng_gens = df_gens[df_gens['name'].str.contains('LNG', na=False)]

print("\n석탄발전기:")
print(f"  수: {len(coal_gens)}")
print(f"  총 용량: {coal_gens['p_nom'].sum():.1f} MW")
print(f"  평균 효율: {coal_gens['efficiency'].mean():.3f}")
print(f"  효율 범위: {coal_gens['efficiency'].min():.3f} ~ {coal_gens['efficiency'].max():.3f}")

print("\nLNG발전기:")
print(f"  수: {len(lng_gens)}")
print(f"  총 용량: {lng_gens['p_nom'].sum():.1f} MW")
print(f"  평균 효율: {lng_gens['efficiency'].mean():.3f}")

# 4. 석탄의 CO2 배출계수 추정 (일반적으로 약 0.95 tCO2/MWh)
print("\n" + "=" * 80)
print("잠재적 CO2 배출량 추정")
print("=" * 80)

coal_emission_factor = 0.95  # tCO2/MWh (석탄)
lng_emission_factor = 0.35   # tCO2/MWh (LNG)

print(f"\n가정:")
print(f"  석탄 배출계수: {coal_emission_factor} tCO2/MWh")
print(f"  LNG 배출계수: {lng_emission_factor} tCO2/MWh")

print(f"\n만약 모든 발전기를 100% 가동한다면:")
total_hours = 8760  # 1년
coal_max_emission = coal_gens['p_nom'].sum() * total_hours * coal_emission_factor / 1000  # 천톤
lng_max_emission = lng_gens['p_nom'].sum() * total_hours * lng_emission_factor / 1000

print(f"  석탄 최대 배출량: {coal_max_emission:.1f} 천톤CO2/년")
print(f"  LNG 최대 배출량: {lng_max_emission:.1f} 천톤CO2/년")
print(f"  합계: {coal_max_emission + lng_max_emission:.1f} 천톤CO2/년")
