"""패턴 로딩 테스트"""
import sys
import os

# src 경로 추가
sys.path.insert(0, os.path.join(os.getcwd(), 'src'))

from process_regional_excel import RegionalExcelProcessor
import pandas as pd

print("=" * 80)
print("패턴 로딩 테스트")
print("=" * 80)

# 1. integrated_input_data.xlsx 생성
print("\n[1] integrated_input_data.xlsx 생성 중...")
processor = RegionalExcelProcessor('interface.xlsx')

if processor.read_data():
    processor.create_integrated_data()
    print("  → 생성 완료!")
else:
    print("  → 생성 실패!")
    sys.exit(1)

# 2. load_patterns 확인
print("\n[2] Load Patterns 확인:")
try:
    lp_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='load_patterns')
    print(f"  행 수: {len(lp_df)}")
    print(f"  열 수: {len(lp_df.columns)}")
    print(f"\n  열 이름 (처음 10개):")
    for i, col in enumerate(lp_df.columns[:10]):
        print(f"    {i+1:2d}. '{col}'")
    
    # 전력 패턴 확인
    if '전력(electricity)' in lp_df.columns or 'electricity' in str(lp_df.columns):
        print(f"\n  ✓ 전력 패턴 열 발견!")
    else:
        print(f"\n  ✗ 전력 패턴 열 없음")
        
except Exception as e:
    print(f"  오류: {e}")

# 3. renewable_patterns 확인
print("\n[3] Renewable Patterns 확인:")
try:
    rp_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='renewable_patterns')
    print(f"  행 수: {len(rp_df)}")
    print(f"  열 수: {len(rp_df.columns)}")
    print(f"\n  열 이름 (처음 10개):")
    for i, col in enumerate(rp_df.columns[:10]):
        print(f"    {i+1:2d}. '{col}'")
    
    # PV 패턴 확인
    pv_patterns = [col for col in rp_df.columns if 'PV_Patterns' in col]
    wt_patterns = [col for col in rp_df.columns if 'WT_Patterns' in col]
    
    print(f"\n  PV 패턴 수: {len(pv_patterns)}개")
    if pv_patterns:
        print(f"    샘플: {pv_patterns[:3]}")
    
    print(f"  WT 패턴 수: {len(wt_patterns)}개")
    if wt_patterns:
        print(f"    샘플: {wt_patterns}")
    
    if pv_patterns:
        # 첫 PV 패턴의 값 확인
        first_pv = pv_patterns[0]
        values = rp_df[first_pv].dropna()
        if len(values) > 0:
            print(f"\n  {first_pv} 값 샘플:")
            print(f"    최소: {values.min():.6f}")
            print(f"    최대: {values.max():.6f}")
            print(f"    평균: {values.mean():.6f}")
            print(f"    데이터 수: {len(values)}")
        
except Exception as e:
    print(f"  오류: {e}")

print("\n" + "=" * 80)
print("결론:")
if 'pv_patterns' in locals() and len(pv_patterns) > 0:
    print("  ✓ 패턴이 제대로 로드되었습니다!")
else:
    print("  ✗ 패턴 로드 실패!")
print("=" * 80)
