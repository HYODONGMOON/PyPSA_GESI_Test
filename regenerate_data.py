"""integrated_input_data.xlsx 재생성"""
import sys
import os

# PyPSA_GUI에서 ensure_integrated_input 함수 import
sys.path.insert(0, os.getcwd())

from PyPSA_GUI import ensure_integrated_input

print("=" * 80)
print("integrated_input_data.xlsx 재생성")
print("=" * 80)

# 기존 파일 삭제
if os.path.exists('integrated_input_data.xlsx'):
    os.remove('integrated_input_data.xlsx')
    print("기존 파일 삭제 완료")

# 재생성
print("\n데이터 생성 중...")
ensure_integrated_input()

if os.path.exists('integrated_input_data.xlsx'):
    print("\n✓ integrated_input_data.xlsx 생성 완료!")
else:
    print("\n✗ 생성 실패!")
    sys.exit(1)

# 패턴 확인
import pandas as pd

print("\n" + "=" * 80)
print("생성된 패턴 확인")
print("=" * 80)

# Load patterns
try:
    lp_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='load_patterns')
    print(f"\nLoad Patterns:")
    print(f"  열 수: {len(lp_df.columns)}")
    
    elec_cols = [col for col in lp_df.columns if 'electricity' in str(col).lower() or '전력' in str(col)]
    print(f"  전력 패턴: {len(elec_cols)}개")
    if elec_cols:
        print(f"    → {elec_cols[:3]}")
except Exception as e:
    print(f"  Load patterns 오류: {e}")

# Renewable patterns
try:
    rp_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='renewable_patterns')
    print(f"\nRenewable Patterns:")
    print(f"  열 수: {len(rp_df.columns)}")
    
    pv_cols = [col for col in rp_df.columns if 'PV_Patterns' in str(col)]
    wt_cols = [col for col in rp_df.columns if 'WT_Patterns' in str(col)]
    
    print(f"  PV 패턴: {len(pv_cols)}개")
    if pv_cols:
        print(f"    → {pv_cols[:3]}")
    print(f"  WT 패턴: {len(wt_cols)}개")
    if wt_cols:
        print(f"    → {wt_cols}")
    
    if pv_cols:
        # 값 확인
        first_pv = pv_cols[0]
        values = rp_df[first_pv].dropna()
        if len(values) > 0:
            print(f"\n  {first_pv} 통계:")
            print(f"    데이터 수: {len(values)}")
            print(f"    최소: {values.min():.6f}")
            print(f"    최대: {values.max():.6f}")
            print(f"    평균: {values.mean():.6f}")
        
except Exception as e:
    print(f"  Renewable patterns 오류: {e}")

print("\n" + "=" * 80)
