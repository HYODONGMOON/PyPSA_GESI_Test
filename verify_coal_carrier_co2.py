import pandas as pd

print("=" * 80)
print("석탄발전기 carrier 및 CO2 배출 설정 확인")
print("=" * 80)

df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

# 석탄발전기 필터링
coal_gens = df_gens[df_gens['name'].str.contains('당진|영흥|하동|삼척|태안|보령|삼천포|호남|동해', na=False)]

print(f"\n석탄발전기 수: {len(coal_gens)}")
print(f"총 용량: {coal_gens['p_nom'].sum():.1f} MW")

# Carrier 확인
print(f"\ncarrier 분포:")
carrier_counts = coal_gens['carrier'].value_counts()
for carrier, count in carrier_counts.items():
    print(f"  {carrier}: {count}개")

if 'coal' in carrier_counts.index:
    print(f"\n[OK] 석탄발전기의 carrier가 'coal'로 설정되었습니다!")
else:
    print(f"\n[경고] 석탄발전기의 carrier가 'coal'로 설정되지 않았습니다!")

# 샘플 발전기 확인
print(f"\n샘플 석탄발전기 (상위 5개):")
print(coal_gens[['name', 'bus', 'carrier', 'p_nom', 'efficiency', 'marginal_cost']].head(5))

# LNG와 비교
lng_gens = df_gens[df_gens['name'].str.contains('LNG', na=False)]
print(f"\n" + "=" * 80)
print("비교: LNG 발전기")
print("=" * 80)
print(f"\nLNG 발전기 carrier 분포:")
lng_carrier_counts = lng_gens['carrier'].value_counts()
for carrier, count in lng_carrier_counts.items():
    print(f"  {carrier}: {count}개")

# Carriers 시트 확인 (CO2 배출계수 설정)
print("\n" + "=" * 80)
print("Carriers 시트 확인 (CO2 배출계수)")
print("=" * 80)

try:
    df_carriers = pd.read_excel('integrated_input_data.xlsx', sheet_name='carriers')
    print(f"\n[OK] carriers 시트 존재")
    print(f"컬럼: {df_carriers.columns.tolist()}")
    print(f"\ncarriers 데이터:")
    print(df_carriers)
    
    # coal carrier 확인
    if 'coal' in df_carriers['name'].values:
        coal_carrier = df_carriers[df_carriers['name'] == 'coal']
        print(f"\n[OK] 'coal' carrier 설정 확인:")
        print(coal_carrier)
    else:
        print(f"\n[경고] 'coal' carrier가 carriers 시트에 정의되지 않았습니다!")
        
except Exception as e:
    print(f"\n[경고] carriers 시트 읽기 실패: {e}")
    print("carriers 시트가 없으면 PyPSA에서 CO2 배출량을 계산하지 않습니다.")
    print("interface.xlsx에 carriers 시트를 추가해야 합니다.")

# Interface.xlsx에서 carriers 확인
print("\n" + "=" * 80)
print("interface.xlsx의 carriers 시트 확인")
print("=" * 80)

try:
    df_interface_carriers = pd.read_excel('interface.xlsx', sheet_name='carriers')
    print(f"\n[OK] interface.xlsx에 carriers 시트 존재")
    print(f"컬럼: {df_interface_carriers.columns.tolist()}")
    print(f"\ncarriers 데이터:")
    print(df_interface_carriers)
    
    # coal carrier 확인
    if 'coal' in df_interface_carriers['name'].values:
        coal_carrier = df_interface_carriers[df_interface_carriers['name'] == 'coal']
        print(f"\n[OK] 'coal' carrier CO2 배출계수:")
        print(coal_carrier)
    else:
        print(f"\n[경고] 'coal' carrier가 정의되지 않았습니다!")
        
except Exception as e:
    print(f"\n[경고] interface.xlsx의 carriers 시트 읽기 실패: {e}")
    print("\n[권장] interface.xlsx에 다음과 같은 carriers 시트를 추가하세요:")
    print("\nname       | co2_emissions (tCO2/MWh_th)")
    print("-" * 40)
    print("coal       | 0.340")
    print("gas        | 0.198")
    print("nuclear    | 0.000")
    print("solar      | 0.000")
    print("wind       | 0.000")
    print("electricity| 0.000")
