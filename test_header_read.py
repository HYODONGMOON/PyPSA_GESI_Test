import pandas as pd

print("=" * 80)
print("Header 읽기 테스트")
print("=" * 80)

# 1. header=3으로 읽기
print("\n[1] header=3으로 읽기:")
df_h3 = pd.read_excel('interface.xlsx', sheet_name='renewable_patterns', header=3)
print(f"  행 수: {len(df_h3)}")
print(f"  열 이름 (처음 5개): {df_h3.columns[:5].tolist()}")
print(f"\n  첫 3행:")
print(df_h3.head(3))

# 2. header=None으로 읽고 수동 처리
print("\n[2] header=None으로 읽고 수동 처리:")
df_raw = pd.read_excel('interface.xlsx', sheet_name='renewable_patterns', header=None)
print(f"  원본 행 수: {len(df_raw)}")
print(f"\n  행 3 (헤더):")
print(df_raw.iloc[3])

# 헤더를 행 3으로 설정하고 행 4부터를 데이터로
df_manual = df_raw.iloc[4:].copy()
df_manual.columns = df_raw.iloc[3].tolist()
df_manual.reset_index(drop=True, inplace=True)

print(f"\n  처리 후 행 수: {len(df_manual)}")
print(f"  열 이름 (처음 5개): {df_manual.columns[:5].tolist()}")
print(f"\n  첫 3행:")
print(df_manual.head(3))

# 3. skiprows로 읽기
print("\n[3] skiprows=[0,1,2]로 읽기:")
df_skip = pd.read_excel('interface.xlsx', sheet_name='renewable_patterns', skiprows=[0,1,2])
print(f"  행 수: {len(df_skip)}")
print(f"  열 이름 (처음 5개): {df_skip.columns[:5].tolist()}")
print(f"\n  첫 3행:")
print(df_skip.head(3))

print("\n" + "=" * 80)
print("결론:")
print("=" * 80)
if len(df_skip) == 8760 and 'BSN_PV_Patterns' in df_skip.columns:
    print("✓ skiprows 방식이 올바릅니다!")
elif len(df_manual) == 8760 and 'BSN_PV_Patterns' in df_manual.columns:
    print("✓ 수동 처리가 올바릅니다!")
else:
    print("✗ 모든 방법이 실패했습니다!")
print("=" * 80)
