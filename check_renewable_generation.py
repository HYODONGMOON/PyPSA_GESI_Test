import pandas as pd
import numpy as np
from PyPSA_GUI import read_input_data, create_network, standardize_bus_names_in_input

# 데이터 로드 및 네트워크 생성
input_data = read_input_data('integrated_input_data.xlsx')
input_data = standardize_bus_names_in_input(input_data)
network = create_network(input_data)

print('='*70)
print('재생에너지 발전 패턴 검증')
print('='*70)

# 1. 재생에너지 발전기 확인
re_gens = network.generators[network.generators['carrier'].isin(['solar', 'wind'])]
print(f'\n1. 재생에너지 발전기:')
print(f'   - 총 {len(re_gens)}개')
print(f'   - Solar: {len(re_gens[re_gens["carrier"]=="solar"])}개')
print(f'   - Wind: {len(re_gens[re_gens["carrier"]=="wind"])}개')
print(f'   - 총 설비용량: {re_gens["p_nom"].sum():.2f} MW')

# 2. generators_t.p_max_pu 확인
if hasattr(network.generators_t, 'p_max_pu') and not network.generators_t.p_max_pu.empty:
    print(f'\n2. generators_t.p_max_pu:')
    print(f'   - Shape: {network.generators_t.p_max_pu.shape}')
    
    # 재생에너지 발전기의 p_max_pu
    re_gen_names = re_gens.index.tolist()
    re_cols = [col for col in network.generators_t.p_max_pu.columns if col in re_gen_names]
    
    print(f'   - 재생에너지 발전기 중 p_max_pu 있음: {len(re_cols)}개')
    
    if len(re_cols) > 0:
        re_pmax = network.generators_t.p_max_pu[re_cols]
        print(f'   - 전체 기간 평균: {re_pmax.mean().mean():.4f}')
        print(f'   - 구간 1 (732시간) 평균: {re_pmax.iloc[:732].mean().mean():.4f}')
        
        # 샘플: BSN_PV
        bsn_pv_cols = [col for col in re_cols if col.startswith('BSN_') and 'PV' in col]
        if len(bsn_pv_cols) > 0:
            sample_col = bsn_pv_cols[0]
            sample_data = network.generators_t.p_max_pu[sample_col]
            print(f'\n3. 샘플: {sample_col}')
            print(f'   - 평균: {sample_data.mean():.4f}')
            print(f'   - 최대: {sample_data.max():.4f}')
            print(f'   - 최소: {sample_data.min():.4f}')
            print(f'   - 처음 24시간 평균: {sample_data.iloc[:24].mean():.4f}')
            print(f'   - 처음 5개 값: {sample_data.iloc[:5].tolist()}')
        
        # 구간 1 재생에너지 잠재 발전량
        seg1_re_potential = 0.0
        for col in re_cols:
            gen_name = col
            if gen_name in network.generators.index:
                p_nom = network.generators.at[gen_name, 'p_nom']
                p_max_pu_seg1 = network.generators_t.p_max_pu[col].iloc[:732]
                seg1_potential = (p_nom * p_max_pu_seg1).sum()
                seg1_re_potential += seg1_potential
        
        print(f'\n4. 구간 1 재생에너지 잠재 발전량:')
        print(f'   - 총 잠재 발전량: {seg1_re_potential:.2f} MWh = {seg1_re_potential/1e6:.2f} GWh')
        print(f'   - 구간 1 수요: 71.69 GWh')
        print(f'   - 재생에너지 비율: {seg1_re_potential/71.69e6*100:.2f}%')
else:
    print('\n2. generators_t.p_max_pu가 비어있습니다!')

# 5. 전체 발전 설비용량
print(f'\n5. 전체 발전 설비:')
print(f'   - 총 발전기: {len(network.generators)}개')
print(f'   - 총 설비용량: {network.generators["p_nom"].sum():.2f} MW')

print('\n' + '='*70)
