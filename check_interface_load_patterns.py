import pandas as pd

print("=" * 80)
print("interface.xlsx의 load_patterns 시트 확인")
print("=" * 80)

try:
    lp_df = pd.read_excel('interface.xlsx', sheet_name='load_patterns')
    
    print(f"\n기본 정보:")
    print(f"  행 수: {len(lp_df)}")
    print(f"  열 수: {len(lp_df.columns)}")
    
    print(f"\n열 이름 (처음 20개):")
    for i, col in enumerate(lp_df.columns[:20]):
        print(f"    {i+1:3d}. '{col}'")
    
    if len(lp_df.columns) > 20:
        print(f"    ... ({len(lp_df.columns) - 20}개 더 있음)")
    
    print(f"\n첫 5행:")
    print(lp_df.head().to_string())
    
    # 부하 이름이 있는지 확인
    expected_loads = ['BSN_Demand_EL', 'CBD_Demand_EL', 'CND_Demand_EL', 'DGU_Demand_EL']
    print(f"\n예상 부하 이름 확인:")
    for load_name in expected_loads:
        if load_name in lp_df.columns:
            print(f"  OK: {load_name}")
        else:
            print(f"  MISS: {load_name}")
    
    # 시간 열 확인
    if '시간' in lp_df.columns or 'hour' in lp_df.columns:
        print(f"\n시간 열 찾음!")
    else:
        print(f"\n시간 열 없음!")
        print(f"  첫 번째 열 이름: '{lp_df.columns[0]}'")
        print(f"  첫 번째 열 값 샘플:")
        print(lp_df.iloc[:10, 0].to_string())
        
except Exception as e:
    print(f"\n오류 발생: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 80)
