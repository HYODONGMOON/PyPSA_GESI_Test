import sys
import pandas as pd
import numpy as np

# PyPSA_GUI의 함수들을 임포트
from PyPSA_GUI import read_input_data, create_network, standardize_bus_names_in_input

print('='*70)
print('네트워크 생성 및 loads_t.p_set 진단')
print('='*70)

# 데이터 로드
print('\n1. 데이터 로드 중...')
input_data = read_input_data('integrated_input_data.xlsx')
input_data = standardize_bus_names_in_input(input_data)

# 네트워크 생성
print('\n2. 네트워크 생성 중...')
network = create_network(input_data)

print(f'\n3. 네트워크 정보:')
print(f'   - Snapshots: {len(network.snapshots)}')
print(f'   - Loads: {len(network.loads)}')
print(f'   - loads_t.p_set shape: {network.loads_t.p_set.shape}')

# loads_t.p_set 분석
print(f'\n4. loads_t.p_set 분석:')
print(f'   - 전체 합계: {network.loads_t.p_set.sum().sum():.2f} MWh')
print(f'   - 평균 시간당: {network.loads_t.p_set.sum(axis=1).mean():.2f} MW')

# EL 부하만 분석
el_cols = [col for col in network.loads_t.p_set.columns if '_Demand_EL' in col]
print(f'\n5. EL 부하 분석:')
print(f'   - EL 부하 수: {len(el_cols)}')
print(f'   - EL 부하 전체 합계: {network.loads_t.p_set[el_cols].sum().sum():.2f} MWh')
print(f'   - EL 부하 평균 시간당: {network.loads_t.p_set[el_cols].sum(axis=1).mean():.2f} MW')

# 구간 1 (732시간) 분석
seg1_hours = 732
print(f'\n6. 구간 1 (처음 {seg1_hours}시간) 분석:')
seg1_demand = network.loads_t.p_set.iloc[:seg1_hours].sum().sum()
print(f'   - 구간 1 전체 수요: {seg1_demand/1e6:.2f} GWh')
seg1_el_demand = network.loads_t.p_set[el_cols].iloc[:seg1_hours].sum().sum()
print(f'   - 구간 1 EL 수요: {seg1_el_demand/1e6:.2f} GWh')

# 샘플 데이터 확인 (BSN_Demand_EL)
if 'BSN_Demand_EL' in network.loads_t.p_set.columns:
    print(f'\n7. BSN_Demand_EL 샘플 (처음 24시간):')
    bsn_data = network.loads_t.p_set['BSN_Demand_EL'].iloc[:24]
    print(f'   - 최소: {bsn_data.min():.2f} MW')
    print(f'   - 최대: {bsn_data.max():.2f} MW')
    print(f'   - 평균: {bsn_data.mean():.2f} MW')
    print(f'   - 처음 5개 값: {bsn_data.iloc[:5].tolist()}')
    
    # 전체 BSN_Demand_EL 통계
    bsn_full = network.loads_t.p_set['BSN_Demand_EL']
    print(f'\n8. BSN_Demand_EL 전체 통계:')
    print(f'   - 연간 총량: {bsn_full.sum():.2f} MWh = {bsn_full.sum()/1e6:.2f} GWh')
    print(f'   - 평균: {bsn_full.mean():.2f} MW')
    print(f'   - 예상 평균 (4623 MW): 4623 MW')
    print(f'   - 비율: {bsn_full.mean() / 4623.0:.4f}')

print('\n' + '='*70)
