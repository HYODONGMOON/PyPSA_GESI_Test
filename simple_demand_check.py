import pandas as pd

print("="*70)
print("간단한 수요 확인")
print("="*70)

# Loads 확인
loads_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='loads')
electric_loads = loads_df[loads_df['bus'].str.contains('_EL', na=False)]

print(f"\n전력 부하 수: {len(electric_loads)}개")
print(f"총 전력 수요(p_set): {electric_loads['p_set'].sum():.0f} MW")

print("\n지역별 수요:")
for bus in sorted(electric_loads['bus'].unique()):
    region_loads = electric_loads[electric_loads['bus'] == bus]
    total = region_loads['p_set'].sum()
    if total > 0:
        print(f"  {bus}: {total:.0f} MW")

# 원자력 비교
nuclear_min = 24650 * 0.8
print(f"\n원자력 최소 출력: {nuclear_min:.0f} MW")
print(f"총 전력 수요: {electric_loads['p_set'].sum():.0f} MW")

if electric_loads['p_set'].sum() > 0:
    ratio = nuclear_min / electric_loads['p_set'].sum()
    print(f"비율: 원자력 최소 / 총 수요 = {ratio:.2f}")
    
    if ratio > 1:
        print(f"\n[문제] 원자력 최소 출력이 총 수요보다 {(ratio-1)*100:.1f}% 높습니다!")
        print(f"야간/저수요 시간대에는 더 큰 문제가 발생합니다.")
else:
    print("\n[경고] 전력 수요가 0입니다. 데이터 확인 필요!")

# Load patterns 시트 확인
try:
    patterns = pd.read_excel('integrated_input_data.xlsx', sheet_name='load_patterns')
    print(f"\n부하 패턴 시트:")
    print(f"  행 수: {len(patterns)}")
    print(f"  열 수: {len(patterns.columns)}")
    
    # 첫 몇 개 열 이름
    print(f"\n  처음 10개 열:")
    for i, col in enumerate(patterns.columns[:10]):
        print(f"    {i+1}. {col}")
        
except Exception as e:
    print(f"\nload_patterns 시트 읽기 오류: {e}")

print("\n"+"="*70)
print("권장 조치:")
print("="*70)
print("p_min_pu를 0.5 이하로 낮추는 것을 강력히 권장합니다.")
print("현재 0.8은 실제 수요 대비 너무 높아 대부분의 시간대에서")
print("infeasible을 유발합니다.")
print("="*70)
