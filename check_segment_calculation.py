import pandas as pd
import numpy as np
from PyPSA_GUI import read_input_data, create_network, standardize_bus_names_in_input

# 데이터 로드 및 네트워크 생성
input_data = read_input_data('integrated_input_data.xlsx')
input_data = standardize_bus_names_in_input(input_data)
network = create_network(input_data)

print('='*70)
print('구간 1 수요 계산 검증')
print('='*70)

# 구간 1 정보
start_hour = 0
end_hour = 732

print(f'\n1. network.loads_t.p_set 전체 정보:')
print(f'   - Shape: {network.loads_t.p_set.shape}')
print(f'   - 전체 합계: {network.loads_t.p_set.sum().sum():.2f} MWh')

# 구간 1 슬라이싱
seg1_data = network.loads_t.p_set.iloc[start_hour:end_hour]
print(f'\n2. 구간 1 슬라이싱 (iloc[{start_hour}:{end_hour}]):')
print(f'   - Shape: {seg1_data.shape}')
print(f'   - 데이터 타입: {type(seg1_data)}')

# 단계별 합계 계산
print(f'\n3. 단계별 합계:')
col_sums = seg1_data.sum()  # 각 열(부하)의 합계
print(f'   - seg1_data.sum() (첫 번째): shape={col_sums.shape}, type={type(col_sums)}')
print(f'   - 처음 5개: {col_sums.iloc[:5].tolist()}')

total_sum = col_sums.sum()  # 전체 합계
print(f'   - col_sums.sum() (두 번째): {total_sum:.2f}')

# rolling_horizon_core의 계산 재현
segment_demand = float(network.loads_t.p_set.iloc[start_hour:end_hour].sum().sum())
print(f'\n4. rolling_horizon_core 계산 재현:')
print(f'   - segment_demand = {segment_demand:.2f} MWh')
print(f'   - segment_demand / 1e6 = {segment_demand/1e6:.2f} GWh')
print(f'   - segment_demand / 1e3 = {segment_demand/1e3:.2f} GWh')

# 예상 값과 비교
expected = 75776.93 * 732
print(f'\n5. 예상 값과 비교:')
print(f'   - 예상: {expected:.2f} MWh = {expected/1e6:.2f} GWh')
print(f'   - 실제: {segment_demand:.2f} MWh = {segment_demand/1e6:.2f} GWh')
print(f'   - 비율: {segment_demand / expected:.6f}')

print('\n' + '='*70)
