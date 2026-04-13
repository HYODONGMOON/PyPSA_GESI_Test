import pandas as pd

print("=" * 70)
print("원자력 발전소 제약 조건 확인")
print("=" * 70)

# integrated_input_data.xlsx 읽기
try:
    generators_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
    
    # 원자력 발전소 필터링
    nuclear_gens = generators_df[generators_df['carrier'].str.contains('원자력|Nuclear|nuclear', case=False, na=False)]
    
    if len(nuclear_gens) == 0:
        # carrier로 찾지 못한 경우 name으로 검색
        nuclear_gens = generators_df[generators_df['name'].str.contains('원자력|Nuclear|nuclear', case=False, na=False)]
    
    print(f"\n원자력 발전소 수: {len(nuclear_gens)}개")
    
    if len(nuclear_gens) > 0:
        print("\n주요 파라미터:")
        cols_to_show = ['name', 'bus', 'carrier', 'p_nom', 'p_min_pu', 'marginal_cost', 'committable']
        available_cols = [col for col in cols_to_show if col in nuclear_gens.columns]
        
        print(nuclear_gens[available_cols].to_string(index=False))
        
        # p_min_pu 통계
        if 'p_min_pu' in nuclear_gens.columns:
            print(f"\np_min_pu 분포:")
            print(f"  - 최소값: {nuclear_gens['p_min_pu'].min()}")
            print(f"  - 최대값: {nuclear_gens['p_min_pu'].max()}")
            print(f"  - 평균값: {nuclear_gens['p_min_pu'].mean():.3f}")
            
            # 0.8 이상인 발전소
            high_pmin = nuclear_gens[nuclear_gens['p_min_pu'] >= 0.8]
            print(f"\np_min_pu >= 0.8인 원자력 발전소: {len(high_pmin)}개")
            if len(high_pmin) > 0:
                print(f"  총 용량: {high_pmin['p_nom'].sum():.0f} MW")
    else:
        print("\n원자력 발전소를 찾을 수 없습니다.")
        
        # 전체 carrier 종류 확인
        print("\n전체 Carrier 종류:")
        print(generators_df['carrier'].value_counts())
        
except FileNotFoundError:
    print("\nintegrated_input_data.xlsx 파일을 찾을 수 없습니다.")
    print("먼저 데이터를 생성해야 합니다.")
except Exception as e:
    print(f"\n오류 발생: {e}")

print("\n" + "=" * 70)
