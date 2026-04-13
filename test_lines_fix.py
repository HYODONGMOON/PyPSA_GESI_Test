# -*- coding: utf-8 -*-
"""Lines 버스 이름 변환 테스트"""
import sys
import os

# PyPSA_GUI 모듈 import
sys.path.insert(0, os.path.dirname(__file__))
from PyPSA_GUI import ensure_integrated_input, read_input_data, standardize_bus_names_in_input, _persist_standardized_input
import pandas as pd

print("=" * 70)
print("Lines 버스 이름 변환 테스트")
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
print("\n[4단계] Lines 버스 이름 검증...")
print("=" * 70)

if 'lines' in input_data and not input_data['lines'].empty:
    lines_df = input_data['lines']
    print(f"\n[OK] Lines: {len(lines_df)}개")
    
    # 상위 20개
    print("\n상위 20개 lines:")
    print(lines_df[['name', 'bus0', 'bus1']].head(20).to_string(index=False))
    
    # Buses 로드
    buses_df = input_data['buses']
    bus_names = set(buses_df['name'].tolist())
    
    # bus0/bus1이 buses에 있는지 확인
    missing_bus0 = lines_df[~lines_df['bus0'].isin(bus_names)]
    missing_bus1 = lines_df[~lines_df['bus1'].isin(bus_names)]
    
    if len(missing_bus0) > 0:
        print(f"\n[경고] bus0가 buses에 없는 lines: {len(missing_bus0)}개")
        print(missing_bus0[['name', 'bus0', 'bus1']].head(10).to_string(index=False))
    else:
        print("\n[OK] 모든 bus0가 buses에 존재합니다!")
    
    if len(missing_bus1) > 0:
        print(f"\n[경고] bus1이 buses에 없는 lines: {len(missing_bus1)}개")
        print(missing_bus1[['name', 'bus0', 'bus1']].head(10).to_string(index=False))
    else:
        print("\n[OK] 모든 bus1이 buses에 존재합니다!")
else:
    print("\n[실패] Lines 시트가 비어있습니다!")

print("\n" + "=" * 70)
print("테스트 완료!")
print("=" * 70)
