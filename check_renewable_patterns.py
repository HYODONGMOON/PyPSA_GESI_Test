import pandas as pd

print("=" * 80)
print("재생에너지 패턴 확인")
print("=" * 80)

# 1. interface.xlsx의 renewable_patterns 확인
print("\n[1] interface.xlsx의 renewable_patterns 시트:")
try:
    interface_rp = pd.read_excel('interface.xlsx', sheet_name='renewable_patterns')
    print(f"  행 수: {len(interface_rp)}")
    print(f"  열 수: {len(interface_rp.columns)}")
    print(f"\n  열 이름 (처음 20개):")
    for i, col in enumerate(interface_rp.columns[:20]):
        print(f"    {i+1:3d}. '{col}'")
    
    # 지역 코드가 열 이름에 있는지 확인
    regions = ['BSN', 'CBD', 'CND', 'DGU', 'GGD', 'SEL']
    print(f"\n  지역별 재생에너지 패턴 확인:")
    for region in regions:
        # 해당 지역이 포함된 열 찾기
        matching_cols = [col for col in interface_rp.columns if region in str(col)]
        if matching_cols:
            print(f"    {region}: {len(matching_cols)}개 열 발견")
            for col in matching_cols[:3]:
                print(f"      - {col}")
        else:
            print(f"    {region}: 없음")
    
    print(f"\n  첫 5행:")
    print(interface_rp.head().to_string())
    
except Exception as e:
    print(f"  오류: {e}")

# 2. integrated_input_data.xlsx의 renewable_patterns 확인
print("\n" + "=" * 80)
print("[2] integrated_input_data.xlsx의 renewable_patterns 시트:")
try:
    integrated_rp = pd.read_excel('integrated_input_data.xlsx', sheet_name='renewable_patterns')
    print(f"  행 수: {len(integrated_rp)}")
    print(f"  열 수: {len(integrated_rp.columns)}")
    print(f"\n  열 이름 (처음 20개):")
    for i, col in enumerate(integrated_rp.columns[:20]):
        print(f"    {i+1:3d}. '{col}'")
    
    # 발전기 이름이 있는지 확인
    print(f"\n  발전기 이름 패턴 확인:")
    sample_gen_names = ['BSN_Solar', 'BSN_Wind_onshore', 'GGD_Solar', 'SEL_Solar']
    for gen_name in sample_gen_names:
        if gen_name in integrated_rp.columns:
            # 처음 몇 개 값 확인
            values = integrated_rp[gen_name].dropna().head(10).values
            if len(values) > 0:
                print(f"    OK: {gen_name} (첫 값: {values[0]:.6f})")
            else:
                print(f"    EMPTY: {gen_name}")
        else:
            print(f"    MISS: {gen_name}")
    
    print(f"\n  첫 5행 (처음 10개 열만):")
    print(integrated_rp.iloc[:5, :min(10, len(integrated_rp.columns))].to_string())
    
except Exception as e:
    print(f"  오류: {e}")

# 3. generators 확인
print("\n" + "=" * 80)
print("[3] Generators의 재생에너지 발전기 확인:")
try:
    gens_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
    
    # 재생에너지 필터링
    re_types = ['Solar', 'Wind', 'solar', 'wind', 'PV']
    re_gens = gens_df[gens_df['name'].str.contains('|'.join(re_types), case=False, na=False)]
    
    print(f"  재생에너지 발전기 수: {len(re_gens)}개")
    
    if len(re_gens) > 0:
        print(f"\n  샘플 (처음 10개):")
        cols_to_show = ['name', 'bus', 'carrier', 'p_nom']
        print(re_gens[cols_to_show].head(10).to_string(index=False))
        
        # renewable_patterns에 이름이 있는지 확인
        if 'integrated_rp' in locals():
            print(f"\n  renewable_patterns와 매칭 확인:")
            matched = 0
            for gen_name in re_gens['name'].head(10):
                if gen_name in integrated_rp.columns:
                    print(f"    OK: {gen_name}")
                    matched += 1
                else:
                    print(f"    MISS: {gen_name}")
            
            print(f"\n  매칭 비율: {matched}/10")
    
except Exception as e:
    print(f"  오류: {e}")

print("\n" + "=" * 80)
print("결론:")
print("=" * 80)
print("재생에너지 패턴이 제대로 로드되지 않았다면,")
print("generators_t.p_max_pu가 채워지지 않아 재생에너지 발전이 0이 됩니다.")
print("이것도 infeasible의 원인이 될 수 있습니다!")
print("=" * 80)
