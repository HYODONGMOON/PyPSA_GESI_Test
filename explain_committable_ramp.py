import pandas as pd

print("=" * 80)
print("Committable 옵션과 Ramp 제약 설명")
print("=" * 80)

df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

# 석탄과 LNG 발전기 필터링
coal_gens = df_gens[df_gens['carrier'] == 'coal']
lng_gens = df_gens[df_gens['name'].str.contains('LNG', na=False)]

print("\n" + "=" * 80)
print("1. 현재 설정 확인")
print("=" * 80)

print("\n[석탄발전기]")
print(f"  발전기 수: {len(coal_gens)}")
if len(coal_gens) > 0:
    print(f"  committable 설정:")
    committable_counts = coal_gens['committable'].value_counts()
    for val, count in committable_counts.items():
        print(f"    {val}: {count}개")
    
    print(f"  p_min_pu 설정:")
    p_min_pu_unique = coal_gens['p_min_pu'].unique()
    for val in p_min_pu_unique:
        count = len(coal_gens[coal_gens['p_min_pu'] == val])
        print(f"    {val}: {count}개")
    
    print(f"  ramp_limit_up 설정:")
    ramp_up_unique = coal_gens['ramp_limit_up'].unique()
    for val in ramp_up_unique:
        count = len(coal_gens[coal_gens['ramp_limit_up'] == val])
        print(f"    {val}: {count}개")
    
    print(f"  ramp_limit_down 설정:")
    ramp_down_unique = coal_gens['ramp_limit_down'].unique()
    for val in ramp_down_unique:
        count = len(coal_gens[coal_gens['ramp_limit_down'] == val])
        print(f"    {val}: {count}개")

print("\n[LNG발전기]")
print(f"  발전기 수: {len(lng_gens)}")
if len(lng_gens) > 0:
    print(f"  committable 설정:")
    committable_counts = lng_gens['committable'].value_counts()
    for val, count in committable_counts.items():
        print(f"    {val}: {count}개")
    
    print(f"  p_min_pu 설정:")
    p_min_pu_unique = lng_gens['p_min_pu'].unique()
    for val in p_min_pu_unique:
        count = len(lng_gens[lng_gens['p_min_pu'] == val])
        print(f"    {val}: {count}개")
    
    print(f"  ramp_limit_up 설정:")
    ramp_up_unique = lng_gens['ramp_limit_up'].unique()
    for val in ramp_up_unique:
        count = len(lng_gens[lng_gens['ramp_limit_up'] == val])
        print(f"    {val}: {count}개")

print("\n" + "=" * 80)
print("2. Committable 옵션 설명")
print("=" * 80)

print("""
committable 옵션은 발전기의 '기동정지 최적화(Unit Commitment)'를 활성화합니다.

[committable = False] (기본값, 일반 경제급전)
  - 발전기는 0 ~ p_nom 범위에서 자유롭게 출력 조절 가능
  - p_min_pu는 무시됨 (0부터 출력 가능)
  - 발전기를 켜고 끄는 개념이 없음
  - 단순히 marginal_cost에 따라 경제급전만 수행
  
[committable = True] (Unit Commitment 최적화)
  - 발전기의 ON/OFF 상태를 이진 변수로 모델링
  - ON 상태일 때: p_min_pu × p_nom ~ p_nom 범위에서 출력
  - OFF 상태일 때: 출력 = 0
  - 추가 제약 적용:
    * p_min_pu: 최소 출력 (ON 상태일 때 반드시 이 이상 출력해야 함)
    * min_up_time: 최소 가동 시간 (켜면 최소 몇 시간 이상 가동)
    * min_down_time: 최소 정지 시간 (끄면 최소 몇 시간 이상 정지)
    * start_up_cost: 기동 비용
    * shut_down_cost: 정지 비용
  
[주의사항]
  - committable=True를 많은 발전기에 적용하면 문제가 매우 복잡해짐
    → Mixed Integer Linear Programming (MILP)로 변환되어 풀이 시간 증가
    → Infeasible 발생 가능성 증가 (제약이 너무 많아서)
  - 보통 큰 화력발전소(석탄, 원자력)에만 적용하고, 
    소규모 발전기나 재생에너지는 False로 유지
""")

print("\n" + "=" * 80)
print("3. Ramp Limit 설명")
print("=" * 80)

print("""
ramp_limit_up / ramp_limit_down은 committable과 독립적으로 작동합니다!

[ramp_limit_up] (시간당 최대 출력 증가율)
  - 단위: p_nom 대비 비율 (per hour)
  - 예: ramp_limit_up = 0.125 → 시간당 최대 12.5% 증가 가능
  - 1000MW 발전기 → 시간당 최대 125MW 증가 가능
  
[ramp_limit_down] (시간당 최대 출력 감소율)
  - 단위: p_nom 대비 비율 (per hour)
  - 예: ramp_limit_down = 0.125 → 시간당 최대 12.5% 감소 가능
  - 1000MW 발전기 → 시간당 최대 125MW 감소 가능

[적용 방식]
  - committable = False인 경우에도 ramp 제약은 적용됨!
  - 예: 현재 500MW 출력 중
    * ramp_limit_up = 0.125, p_nom = 1000MW
    * 다음 시간 최대 출력: 500 + 1000×0.125 = 625MW
    
[실제 효과]
  - 발전기가 급격하게 출력을 변경하지 못하도록 제한
  - 재생에너지 변동성에 대응할 수 있는 유연성 제약
  - 석탄발전소는 보통 0.10~0.15 (10~15%/시간)
  - LNG는 더 빠름: 0.20~0.30 (20~30%/시간)
  - 재생에너지는 보통 ramp 제약 없음 (또는 매우 큼)
""")

print("\n" + "=" * 80)
print("4. 조합 시나리오")
print("=" * 80)

print("""
[시나리오 1] committable=False, ramp_limit=0.125
  → 일반 경제급전 + 출력 변화율 제한
  → 0~p_nom 범위에서 자유롭게 출력하되, 급격한 변화는 불가
  → 가장 간단하고 안정적인 설정
  
[시나리오 2] committable=True, p_min_pu=0.4, ramp_limit=0.125
  → Unit Commitment + 최소출력 40% + 출력 변화율 제한
  → ON 상태: 40%~100% 출력, OFF 상태: 0%
  → ON/OFF 전환 최적화 + 출력 변화율 제한
  → 매우 복잡, 큰 발전기에만 적용 권장
  
[시나리오 3] committable=False, ramp_limit=None (또는 매우 큼)
  → 일반 경제급전, 출력 변화 제한 없음
  → 재생에너지나 수력발전에 적합
  
[시나리오 4] committable=True, ramp_limit=None
  → Unit Commitment만 적용, 출력 변화 제한 없음
  → ON/OFF는 최적화하지만, ON 상태에서는 자유롭게 변화
  → 가스터빈 등 빠른 응답이 가능한 발전기에 적합
""")

print("\n" + "=" * 80)
print("5. 현재 설정에 대한 권장사항")
print("=" * 80)

if len(coal_gens) > 0:
    coal_committable = coal_gens['committable'].iloc[0]
    coal_p_min_pu = coal_gens['p_min_pu'].iloc[0]
    coal_ramp_up = coal_gens['ramp_limit_up'].iloc[0]
    
    print(f"\n[현재 석탄발전기 설정]")
    print(f"  committable: {coal_committable}")
    print(f"  p_min_pu: {coal_p_min_pu}")
    print(f"  ramp_limit_up: {coal_ramp_up}")
    
    if coal_committable:
        print(f"\n  [분석] Unit Commitment 활성화")
        print(f"    - {len(coal_gens)}개 발전기에 UC 적용 → 매우 복잡한 MILP 문제")
        print(f"    - p_min_pu={coal_p_min_pu} → ON 상태일 때 최소 {coal_p_min_pu*100:.0f}% 출력")
        print(f"    - ramp_limit={coal_ramp_up} → 시간당 최대 {coal_ramp_up*100:.1f}% 변화")
        print(f"\n  [권장]")
        print(f"    - Infeasible 발생 시: committable=False로 변경 권장")
        print(f"    - 또는 큰 발전기(≥500MW)에만 committable=True 적용")
    else:
        print(f"\n  [분석] 일반 경제급전")
        print(f"    - p_min_pu는 무시됨 (committable=False이므로)")
        print(f"    - ramp_limit={coal_ramp_up} → 시간당 최대 {coal_ramp_up*100:.1f}% 변화")
        print(f"    - 0~100% 범위에서 자유롭게 출력 조절")
        print(f"\n  [권장] 현재 설정이 안정적입니다.")

if len(lng_gens) > 0:
    lng_committable = lng_gens['committable'].iloc[0]
    lng_p_min_pu = lng_gens['p_min_pu'].iloc[0]
    lng_ramp_up = lng_gens['ramp_limit_up'].iloc[0]
    
    print(f"\n[현재 LNG발전기 설정]")
    print(f"  committable: {lng_committable}")
    print(f"  p_min_pu: {lng_p_min_pu}")
    print(f"  ramp_limit_up: {lng_ramp_up}")
    
    if lng_committable:
        print(f"\n  [분석] Unit Commitment 활성화")
        print(f"    - {len(lng_gens)}개 발전기에 UC 적용")
        print(f"    - LNG는 보통 committable=False로 운영 (빠른 응답성)")
        print(f"\n  [권장] committable=False로 변경 고려")
    else:
        print(f"\n  [분석] 일반 경제급전")
        print(f"    - LNG는 marginal_cost가 높아 피크/백업 전원으로 활용")
        print(f"    - 빠른 출력 조절 가능")
        print(f"\n  [권장] 현재 설정이 적절합니다.")
