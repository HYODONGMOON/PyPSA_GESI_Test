import pandas as pd

df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

# 세분화된 Coal 발전기 (발전소 이름 포함)
coal_detailed = df[df['name'].str.contains('당진|영흥|하동|삼척|태안|보령|삼천포|호남|동해', na=False)]

print(f"세분화된 Coal 발전기 수: {len(coal_detailed)}")
print(f"\nCarrier 분포:")
print(coal_detailed['carrier'].value_counts())
print(f"\n샘플:")
print(coal_detailed[['name', 'carrier', 'p_nom', 'bus']].head(15))

# 전체 발전기에서 carrier='coal'인 것이 남아있는지 확인
coal_carrier = df[df['carrier'] == 'coal']
if not coal_carrier.empty:
    print(f"\n[경고] carrier='coal'인 발전기가 {len(coal_carrier)}개 남아있습니다:")
    print(coal_carrier[['name', 'carrier']].head(10))
else:
    print(f"\n[OK] carrier='coal'인 발전기가 없습니다. 모두 'electricity'로 변경되었습니다.")

# carrier='electricity'인 발전기 수
elec_gens = df[df['carrier'] == 'electricity']
print(f"\nCarrier='electricity'인 발전기 수: {len(elec_gens)}")
