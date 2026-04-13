# -*- coding: utf-8 -*-
"""virtual_buses 추가 테스트"""
import pandas as pd
import traceback

try:
    # virtual_buses 로드
    print("[1] virtual_buses 로드...")
    virtual_buses_df = pd.read_excel('interface.xlsx', sheet_name='virtual_buses')
    print(f"[OK] virtual_buses: {len(virtual_buses_df)}개")
    print(f"[OK] Columns: {virtual_buses_df.columns.tolist()}")
    
    # 필요한 컬럼만 추출
    print("\n[2] 필요한 컬럼 추출...")
    virtual_buses_std = virtual_buses_df[['name', 'carrier', 'v_nom', 'x', 'y']].copy()
    print(f"[OK] 추출됨: {len(virtual_buses_std)}개")
    print("\n상위 5개:")
    print(virtual_buses_std.head())
    
except Exception as e:
    print(f"\n[실패] 오류 발생: {str(e)}")
    traceback.print_exc()
