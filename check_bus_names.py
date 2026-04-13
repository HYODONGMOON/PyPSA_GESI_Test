"""Bus 이름 중복 확인"""
import pandas as pd

print("=" * 80)
print("integrated_input_data.xlsx 버스 이름 확인")
print("=" * 80)

# Loads 확인
loads = pd.read_excel('integrated_input_data.xlsx', sheet_name='loads')
print("\n=== Loads (처음 20개) ===")
print(loads[['name', 'bus']].head(20).to_string(index=False))

# 중복 패턴 확인 (지역코드_지역코드_EL)
print("\n=== 중복 패턴 확인 ===")
dup_pattern = loads['bus'].str.match(r'^[A-Z]{3}_[A-Z]{3}_EL$', na=False)
if dup_pattern.any():
    print(f"중복 패턴 발견: {dup_pattern.sum()}개")
    print(loads[dup_pattern][['name', 'bus']].to_string(index=False))
else:
    print("중복 패턴 없음")

# Lines 확인
print("\n" + "=" * 80)
print("Lines 버스 확인")
print("=" * 80)

lines = pd.read_excel('integrated_input_data.xlsx', sheet_name='lines')
print(f"\n총 Lines: {len(lines)}개")
print("\n=== Lines (처음 10개) ===")
print(lines[['name', 'bus0', 'bus1']].head(10).to_string(index=False))

# bus_EL 패턴 확인
bus_el_pattern = (lines['bus0'] == 'bus_EL') | (lines['bus1'] == 'bus_EL')
if bus_el_pattern.any():
    print(f"\n경고: bus_EL 패턴 발견: {bus_el_pattern.sum()}개")
    print(lines[bus_el_pattern][['name', 'bus0', 'bus1']].head(20).to_string(index=False))
else:
    print("\nbus_EL 패턴 없음 (정상)")

# Generators 확인
print("\n" + "=" * 80)
print("Generators 버스 확인")
print("=" * 80)

generators = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
print("\n=== Generators (처음 20개) ===")
print(generators[['name', 'bus']].head(20).to_string(index=False))

# 중복 패턴 확인
dup_pattern_gen = generators['bus'].str.match(r'^[A-Z]{3}_[A-Z]{3}_', na=False)
if dup_pattern_gen.any():
    print(f"\n중복 패턴 발견: {dup_pattern_gen.sum()}개")
    print(generators[dup_pattern_gen][['name', 'bus']].head(10).to_string(index=False))
else:
    print("\n중복 패턴 없음")
