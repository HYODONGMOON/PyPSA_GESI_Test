import pandas as pd
import numpy as np

print("=" * 80)
print("구간별 수요 vs 원자력 최소 출력 분석")
print("=" * 80)

try:
    # Load patterns 읽기
    loads_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='loads')
    load_patterns_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='load_patterns')
    
    print(f"\n부하 데이터: {len(loads_df)}개")
    print(f"시간별 패턴: {len(load_patterns_df)}개 시간")
    
    # 전력 부하만 필터링 (_EL 버스)
    electric_loads = loads_df[loads_df['bus'].str.contains('_EL', na=False)]
    print(f"전력 부하: {len(electric_loads)}개")
    
    # 각 시간대별 총 전력 수요 계산
    total_demand = []
    for idx, row in load_patterns_df.iterrows():
        hour_total = 0
        for load_name in electric_loads['name']:
            if load_name in load_patterns_df.columns:
                p_set = loads_df[loads_df['name'] == load_name]['p_set'].values[0]
                pattern = row[load_name] if load_name in row else 1.0
                hour_total += p_set * pattern
        total_demand.append(hour_total)
    
    total_demand = np.array(total_demand)
    
    print(f"\n전체 8760시간 전력 수요 통계:")
    print(f"  - 최소 수요: {total_demand.min():.0f} MW")
    print(f"  - 최대 수요: {total_demand.max():.0f} MW")
    print(f"  - 평균 수요: {total_demand.mean():.0f} MW")
    
    # 원자력 최소 출력
    nuclear_min = 24650 * 0.8
    print(f"\n원자력 최소 출력: {nuclear_min:.0f} MW")
    
    # 수요가 원자력 최소 출력보다 낮은 시간
    low_demand_hours = total_demand < nuclear_min
    print(f"\n수요 < 원자력 최소 출력인 시간: {low_demand_hours.sum()}시간 ({low_demand_hours.sum()/8760*100:.1f}%)")
    
    if low_demand_hours.sum() > 0:
        print(f"  → 이 시간대들이 infeasible의 주요 원인!")
    
    # 구간별 분석 (48개 구간, 각 183시간)
    print("\n" + "=" * 80)
    print("구간별 상세 분석 (처음 10개 구간)")
    print("=" * 80)
    
    segment_hours = 183
    overlap = 24
    
    for seg_id in range(min(10, 48)):
        start_hour = seg_id * (segment_hours - overlap)
        end_hour = start_hour + segment_hours
        
        if end_hour > 8760:
            end_hour = 8760
        
        segment_demand = total_demand[start_hour:end_hour]
        
        seg_min = segment_demand.min()
        seg_max = segment_demand.max()
        seg_avg = segment_demand.mean()
        
        # 원자력 최소 출력보다 낮은 시간
        low_hours = (segment_demand < nuclear_min).sum()
        
        status = "✓" if seg_id == 2 else "✗"  # 구간 3만 성공
        
        print(f"\n구간 {seg_id+1} {status}:")
        print(f"  시간: {start_hour}~{end_hour}시")
        print(f"  수요: 최소 {seg_min:.0f} / 평균 {seg_avg:.0f} / 최대 {seg_max:.0f} MW")
        print(f"  원자력 최소(19,720 MW) 미만 시간: {low_hours}시간 ({low_hours/segment_hours*100:.1f}%)")
        
        # 수요가 원자력 최소보다 낮은 시간의 평균 부족분
        if low_hours > 0:
            deficit = nuclear_min - segment_demand[segment_demand < nuclear_min]
            print(f"  평균 초과 공급: {deficit.mean():.0f} MW")
            print(f"  최대 초과 공급: {deficit.max():.0f} MW")

    print("\n" + "=" * 80)
    print("결론:")
    print("=" * 80)
    print(f"거의 모든 시간대에서 수요가 원자력 최소 출력(19,720 MW)보다 낮습니다.")
    print(f"이는 p_min_pu=0.8 설정 때문에 발생하는 구조적 문제입니다.")
    print(f"\n해결 방안:")
    print(f"  1. p_min_pu를 0.5 이하로 낮추기 (권장)")
    print(f"  2. committable=True로 설정하여 일부 원자력을 on/off 가능하게")
    print(f"  3. 원자력 용량 자체를 줄이기 (비현실적)")
    
except Exception as e:
    print(f"\n오류 발생: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 80)
