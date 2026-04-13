import pandas as pd

# integrated_input_data.xlsx의 generators 시트 읽기
df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

print(f"총 발전기 수: {len(df_gens)}")

# Coal 발전기만 필터링
coal_gens = df_gens[df_gens['carrier'] == 'coal']
print(f"\nCoal 발전기 수: {len(coal_gens)}")

if not coal_gens.empty:
    print(f"Coal 발전기 총 용량: {coal_gens['p_nom'].sum():.1f} MW")
    
    # 지역별 Coal 발전기 수
    print("\n지역별 Coal 발전기:")
    for region in ['CND', 'GND', 'GWD', 'ICN', 'JND']:
        region_coal = coal_gens[coal_gens['name'].str.startswith(region)]
        if not region_coal.empty:
            print(f"  {region}: {len(region_coal)}개, {region_coal['p_nom'].sum():.1f} MW")
    
    # 샘플 출력
    print("\n개별 Coal 발전기 샘플 (상위 10개):")
    print(coal_gens[['name', 'p_nom', 'efficiency', 'marginal_cost']].head(10))
    
    # 효율 분포 확인
    print(f"\n효율 범위: {coal_gens['efficiency'].min():.3f} ~ {coal_gens['efficiency'].max():.3f}")
    print(f"평균 효율: {coal_gens['efficiency'].mean():.3f}")
else:
    print("Coal 발전기가 없습니다!")
