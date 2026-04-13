"""integrated_input_data.xlsx 재생성 및 검증"""
import sys
sys.path.insert(0, '.')
from PyPSA_GUI import read_input_data, standardize_bus_names_in_input, _persist_standardized_input
import os

print("=" * 80)
print("integrated_input_data.xlsx 재생성")
print("=" * 80)

# 1. interface.xlsx 읽기
print("\n[1단계] interface.xlsx 읽기...")
input_data = read_input_data('interface.xlsx')
if input_data is None:
    print("[실패] interface.xlsx를 읽을 수 없습니다.")
    exit(1)

print(f"[OK] 데이터 로드 완료")
print(f"  - Buses: {len(input_data.get('buses', []))}개")
print(f"  - Generators: {len(input_data.get('generators', []))}개")
print(f"  - Loads: {len(input_data.get('loads', []))}개")
print(f"  - Lines: {len(input_data.get('lines', []))}개")

# 2. 버스명 표준화
print("\n[2단계] 버스명 표준화...")
input_data = standardize_bus_names_in_input(input_data)

# 3. integrated_input_data.xlsx 저장
print("\n[3단계] integrated_input_data.xlsx 저장...")
integrated_path = os.path.abspath('integrated_input_data.xlsx')
_persist_standardized_input(integrated_path, input_data)
print(f"[OK] 저장 완료: {integrated_path}")

# 4. 검증
print("\n[4단계] 버스 이름 검증...")
import pandas as pd

loads = pd.read_excel('integrated_input_data.xlsx', sheet_name='loads')
generators = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
lines = pd.read_excel('integrated_input_data.xlsx', sheet_name='lines')

# 중복 패턴 확인
dup_loads = loads['bus'].str.match(r'^[A-Z]{3}_[A-Z]{3}_EL$', na=False).sum()
dup_gens = generators['bus'].str.match(r'^[A-Z]{3}_[A-Z]{3}_', na=False).sum()
bus_el_lines = ((lines['bus0'] == 'bus_EL') | (lines['bus1'] == 'bus_EL')).sum()

print(f"\n결과:")
print(f"  - Loads 중복 (XXX_XXX_EL): {dup_loads}개 {'❌' if dup_loads > 0 else '✅'}")
print(f"  - Generators 중복: {dup_gens}개 {'❌' if dup_gens > 0 else '✅'}")
print(f"  - Lines bus_EL 패턴: {bus_el_lines}개 {'❌' if bus_el_lines > 0 else '✅'}")

if dup_loads == 0 and dup_gens == 0 and bus_el_lines == 0:
    print("\n[성공] 모든 버스 이름이 정상입니다!")
    print("\n샘플 확인:")
    print("\nLoads:")
    print(loads[['name', 'bus']].head(10).to_string(index=False))
    print("\nLines:")
    print(lines[['name', 'bus0', 'bus1']].head(10).to_string(index=False))
else:
    print("\n[경고] 중복 패턴이 여전히 존재합니다.")
    if dup_loads > 0:
        print("\nLoads 중복 샘플:")
        dup_load_mask = loads['bus'].str.match(r'^[A-Z]{3}_[A-Z]{3}_EL$', na=False)
        print(loads[dup_load_mask][['name', 'bus']].head(5).to_string(index=False))

print("\n" + "=" * 80)
