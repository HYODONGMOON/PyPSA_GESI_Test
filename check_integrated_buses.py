# -*- coding: utf-8 -*-
"""integrated_input_data.xlsx의 buses 확인"""
import pandas as pd

df = pd.read_excel('integrated_input_data.xlsx', sheet_name='buses', engine='openpyxl')

print(f"[OK] buses 시트: 총 {len(df)}개")

# 전력 버스 확인
el_buses = df[df['name'].astype(str).str.endswith('_EL')]
print(f"\n[정보] 전력 버스 (EL): {len(el_buses)}개")

if len(el_buses) > 0:
    print(el_buses[['name', 'carrier']].to_string(index=False))
else:
    print("[경고] 전력 버스가 없습니다!")

# 각 지역별 버스 확인
print("\n[정보] 모든 버스:")
print(df[['name', 'carrier']].to_string(index=False))
