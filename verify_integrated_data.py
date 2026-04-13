import pandas as pd

print("=" * 80)
print("integrated_input_data.xlsx 확인")
print("=" * 80)

df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

print(f"\n전체 발전기 수: {len(df_gens)}")

# 석탄 관련 발전기 확인
coal_gens = df_gens[df_gens['name'].str.contains('Coal|coal|석탄|당진|영흥|하동|삼척|태안|보령|삼천포|호남|동해', case=False, na=False)]

print(f"\n석탄 관련 발전기 수: {len(coal_gens)}")

if len(coal_gens) > 0:
    print("\n발견된 석탄 발전기:")
    print(coal_gens[['name', 'carrier', 'p_nom', 'bus']].head(20))
    
    # Carrier 확인
    print(f"\nCarrier 분포:")
    print(coal_gens['carrier'].value_counts())
else:
    print("\n[ERROR] 석탄 발전기가 integrated_input_data.xlsx에 없습니다!")
    
# 전체 발전기 이름 샘플 확인
print(f"\n전체 발전기 샘플 (상위 30개):")
print(df_gens[['name', 'carrier', 'p_nom']].head(30))

# LNG 발전기 확인
lng_gens = df_gens[df_gens['name'].str.contains('LNG', case=False, na=False)]
print(f"\nLNG 발전기 수: {len(lng_gens)}")
