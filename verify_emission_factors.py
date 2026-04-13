import pandas as pd

print("=" * 80)
print("CO2 배출계수 확인")
print("=" * 80)

# 1. integrated_input_data.xlsx의 carriers 시트 확인
df_carriers = pd.read_excel('integrated_input_data.xlsx', sheet_name='carriers')

print("\n[integrated_input_data.xlsx - carriers 시트]")
print(f"\n컬럼: {df_carriers.columns.tolist()}")
print(f"\nCarriers 데이터:")
print(df_carriers)

# 2. 하드코딩된 값과 비교
print("\n" + "=" * 80)
print("하드코딩된 값과 비교")
print("=" * 80)

hardcoded_coal = 0.8384 * 0.47
hardcoded_gas = 0.38 * 0.53

print(f"\n[하드코딩된 값 (PyPSA_GUI.py)]")
print(f"  coal: 0.8384 × 0.47 = {hardcoded_coal:.6f} tCO2/MWh")
print(f"  gas:  0.38 × 0.53   = {hardcoded_gas:.6f} tCO2/MWh")

coal_value = df_carriers[df_carriers['name'] == 'coal']['co2_emissions'].values[0]
gas_value = df_carriers[df_carriers['name'] == 'gas']['co2_emissions'].values[0]

print(f"\n[carriers 시트의 값]")
print(f"  coal: {coal_value:.6f} tCO2/MWh")
print(f"  gas:  {gas_value:.6f} tCO2/MWh")

# 3. 일치 여부 확인
print(f"\n[일치 여부]")
coal_match = abs(coal_value - hardcoded_coal) < 0.000001
gas_match = abs(gas_value - hardcoded_gas) < 0.000001

if coal_match:
    print(f"  [OK] coal 배출계수 일치!")
else:
    print(f"  [NG] coal 배출계수 불일치! 차이: {abs(coal_value - hardcoded_coal):.6f}")

if gas_match:
    print(f"  [OK] gas(LNG) 배출계수 일치!")
else:
    print(f"  [NG] gas(LNG) 배출계수 불일치! 차이: {abs(gas_value - hardcoded_gas):.6f}")

if coal_match and gas_match:
    print(f"\n[SUCCESS] 모든 배출계수가 하드코딩된 값과 일치합니다!")
else:
    print(f"\n[WARNING] 일부 배출계수가 불일치합니다.")

# 4. 석탄발전기 carrier 재확인
print("\n" + "=" * 80)
print("석탄발전기 carrier 재확인")
print("=" * 80)

df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
coal_gens = df_gens[df_gens['name'].str.contains('당진|영흥|하동|삼척|태안|보령|삼천포|호남|동해', na=False)]

print(f"\n석탄발전기 수: {len(coal_gens)}")
print(f"Carrier 분포:")
carrier_counts = coal_gens['carrier'].value_counts()
for carrier, count in carrier_counts.items():
    print(f"  {carrier}: {count}개")

if 'coal' in carrier_counts.index and carrier_counts['coal'] == len(coal_gens):
    print(f"\n[SUCCESS] 모든 석탄발전기가 carrier='coal'로 설정되었습니다!")
    print(f"  → CO2 배출량이 {coal_value:.6f} tCO2/MWh로 계산됩니다.")
else:
    print(f"\n[WARNING] 일부 석탄발전기의 carrier가 'coal'이 아닙니다!")
