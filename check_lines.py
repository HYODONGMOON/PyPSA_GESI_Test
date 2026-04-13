# -*- coding: utf-8 -*-
"""Lines 시트 확인"""
import pandas as pd

# Lines 시트 로드
df = pd.read_excel('integrated_input_data.xlsx', sheet_name='lines', engine='openpyxl')

print(f"[OK] Lines 시트: 총 {len(df)}개")

# 상위 20개
print("\n상위 20개:")
print(df[['name', 'bus0', 'bus1']].head(20).to_string(index=False))

# bus0와 bus1의 unique 값 확인
print(f"\n\n[정보] bus0 unique: {len(df['bus0'].unique())}개")
print("샘플 bus0:")
print(df['bus0'].unique()[:10])

print(f"\n\n[정보] bus1 unique: {len(df['bus1'].unique())}개")
print("샘플 bus1:")
print(df['bus1'].unique()[:10])

# Buses 시트 로드
buses_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='buses', engine='openpyxl')
bus_names = set(buses_df['name'].tolist())

print(f"\n\n[정보] Buses 시트: 총 {len(buses_df)}개")
print("전력 버스:")
el_buses = [b for b in bus_names if b.endswith('_EL')]
print(el_buses[:20])

# bus_ 버스 확인
bus_pattern = [b for b in bus_names if b.startswith('bus_')]
print(f"\n\n[정보] bus_ 패턴 버스: {len(bus_pattern)}개")
print("샘플:")
print(bus_pattern[:10])

# lines의 bus0/bus1이 buses에 있는지 확인
missing_bus0 = df[~df['bus0'].isin(bus_names)]
missing_bus1 = df[~df['bus1'].isin(bus_names)]

print(f"\n\n[경고] bus0가 buses에 없는 lines: {len(missing_bus0)}개")
if len(missing_bus0) > 0:
    print(missing_bus0[['name', 'bus0', 'bus1']].head(10).to_string(index=False))

print(f"\n\n[경고] bus1이 buses에 없는 lines: {len(missing_bus1)}개")
if len(missing_bus1) > 0:
    print(missing_bus1[['name', 'bus0', 'bus1']].head(10).to_string(index=False))
