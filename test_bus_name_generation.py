# -*- coding: utf-8 -*-
"""버스 이름 생성 로직 테스트"""
import sys
import os

# PyPSA_GUI 모듈 import
sys.path.insert(0, os.path.dirname(__file__))
from PyPSA_GUI import ensure_integrated_input, read_input_data, standardize_bus_names_in_input, _persist_standardized_input
import pandas as pd

print("=" * 70)
print("버스 이름 생성 로직 테스트")
print("=" * 70)

# 1. integrated_input_data.xlsx 생성
print("\n[1단계] integrated_input_data.xlsx 생성...")
ensure_integrated_input()

# 2. 데이터 로드
print("\n[2단계] 데이터 로드...")
input_data = read_input_data('integrated_input_data.xlsx')

# 3. 버스 이름 표준화
print("\n[3단계] 버스 이름 표준화...")
input_data = standardize_bus_names_in_input(input_data)

# 3-1. 표준화된 데이터 저장
print("\n[3-1단계] 표준화된 데이터 저장...")
_persist_standardized_input('integrated_input_data.xlsx', input_data)

# 3-2. 저장된 파일 다시 로드
print("\n[3-2단계] 저장된 파일 다시 로드...")
input_data = read_input_data('integrated_input_data.xlsx')

# 4. 검증
print("\n[4단계] 버스 이름 검증...")
print("=" * 70)

# Buses 시트 확인
if 'buses' in input_data and not input_data['buses'].empty:
    buses_df = input_data['buses']
    print(f"\n[OK] Buses: {len(buses_df)}개")
    print("\n상위 20개 버스:")
    print(buses_df[['name']].head(20).to_string(index=False))
else:
    print("\n[실패] Buses 시트가 비어있습니다!")

# Loads 시트 확인
if 'loads' in input_data and not input_data['loads'].empty:
    loads_df = input_data['loads']
    print(f"\n[OK] Loads: {len(loads_df)}개")
    
    # BSN_BSN_EL 패턴 체크
    duplicated = loads_df[loads_df['bus'].astype(str).str.contains(r'_.*_EL$', regex=True, na=False)]
    if len(duplicated) > 0:
        print(f"\n[경고] 중복 패턴 발견: {len(duplicated)}개")
        print(duplicated[['name', 'bus']].head(10).to_string(index=False))
    else:
        print("\n[OK] 중복 패턴 없음")
    
    print("\n상위 10개 loads:")
    print(loads_df[['name', 'bus']].head(10).to_string(index=False))
else:
    print("\n[실패] Loads 시트가 비어있습니다!")

# Generators 시트 확인
if 'generators' in input_data and not input_data['generators'].empty:
    gens_df = input_data['generators']
    print(f"\n[OK] Generators: {len(gens_df)}개")
    
    # BSN_BSN_EL 패턴 체크
    duplicated = gens_df[gens_df['bus'].astype(str).str.contains(r'_.*_EL$', regex=True, na=False)]
    if len(duplicated) > 0:
        print(f"\n[경고] 중복 패턴 발견: {len(duplicated)}개")
        print(duplicated[['name', 'bus']].head(10).to_string(index=False))
    else:
        print("\n[OK] 중복 패턴 없음")
    
    print("\n상위 10개 generators:")
    print(gens_df[['name', 'bus']].head(10).to_string(index=False))
else:
    print("\n[실패] Generators 시트가 비어있습니다!")

# Lines 시트 확인
if 'lines' in input_data and not input_data['lines'].empty:
    lines_df = input_data['lines']
    print(f"\n[OK] Lines: {len(lines_df)}개")
    
    # bus_EL 패턴 체크
    bus_el_count = len(lines_df[lines_df['bus0'].astype(str) == 'bus_EL'])
    if bus_el_count > 0:
        print(f"\n[경고] bus_EL 패턴 발견: {bus_el_count}개")
        print(lines_df[lines_df['bus0'].astype(str) == 'bus_EL'][['name', 'bus0', 'bus1']].head(10).to_string(index=False))
    else:
        print("\n[OK] bus_EL 패턴 없음")
    
    # bus_로 시작하는 버스 (정상)
    bus_pattern = lines_df[lines_df['bus0'].astype(str).str.startswith('bus_', na=False)]
    print(f"\n[정보] bus_로 시작하는 버스: {len(bus_pattern)}개 (정상)")
    
    print("\n상위 10개 lines:")
    print(lines_df[['name', 'bus0', 'bus1']].head(10).to_string(index=False))
else:
    print("\n[실패] Lines 시트가 비어있습니다!")

print("\n" + "=" * 70)
print("테스트 완료!")
print("=" * 70)
