import pandas as pd

print("=" * 80)
print("Load Patterns 상세 확인")
print("=" * 80)

# Loads 확인
loads_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='loads')
print(f"\nLoads 시트:")
print(f"  총 부하 수: {len(loads_df)}")
print(f"  열: {list(loads_df.columns)}")

electric_loads = loads_df[loads_df['bus'].str.contains('_EL', na=False)]
print(f"\n  전력 부하: {len(electric_loads)}개")
print(f"  총 p_set: {electric_loads['p_set'].sum():.0f} MW")

# Load patterns 확인
patterns_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='load_patterns')
print(f"\nLoad Patterns 시트:")
print(f"  행 수 (시간): {len(patterns_df)}")
print(f"  열 수: {len(patterns_df.columns)}")

print(f"\n  열 이름 (처음 20개):")
for i, col in enumerate(patterns_df.columns[:20]):
    print(f"    {i+1:3d}. '{col}'")

# 전력 부하 이름이 패턴 열에 있는지 확인
print(f"\n전력 부하 이름 매칭 확인:")
matched = 0
for load_name in electric_loads['name'].head(10):
    if load_name in patterns_df.columns:
        print(f"  OK: {load_name}")
        matched += 1
    else:
        print(f"  MISS: {load_name}")

print(f"\n  매칭된 부하: {matched}/{min(10, len(electric_loads))}개")

# 패턴 값 샘플
if len(patterns_df.columns) > 1:
    print(f"\n패턴 값 샘플 (첫 번째 비시간 열):")
    first_data_col = patterns_df.columns[1]
    print(f"  열: {first_data_col}")
    print(f"  처음 10개 값:")
    print(patterns_df[first_data_col].head(10).to_string())
    print(f"  통계:")
    print(f"    min: {patterns_df[first_data_col].min():.6f}")
    print(f"    max: {patterns_df[first_data_col].max():.6f}")
    print(f"    mean: {patterns_df[first_data_col].mean():.6f}")

print("\n" + "=" * 80)
