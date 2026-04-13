import pandas as pd
from PyPSA_GUI import read_input_data, create_network, standardize_bus_names_in_input

# 데이터 로드 및 네트워크 생성
input_data = read_input_data('integrated_input_data.xlsx')
input_data = standardize_bus_names_in_input(input_data)
network = create_network(input_data)

print('='*70)
print('발전기 용량 분석')
print('='*70)

# 1. 슬랙 발전기 확인
slack_gens = network.generators[network.generators.index.str.contains('Slack', case=False, na=False)]
print(f'\n1. 슬랙 발전기:')
print(f'   - 개수: {len(slack_gens)}')
print(f'   - 총 용량: {slack_gens["p_nom"].sum():.2f} MW')
print(f'   - 평균 용량: {slack_gens["p_nom"].mean():.2f} MW')
print(f'   - 최대 용량: {slack_gens["p_nom"].max():.2f} MW')

# 2. 실제 발전기 (슬랙 제외)
real_gens = network.generators[~network.generators.index.str.contains('Slack', case=False, na=False)]
print(f'\n2. 실제 발전기 (슬랙 제외):')
print(f'   - 개수: {len(real_gens)}')
print(f'   - 총 용량: {real_gens["p_nom"].sum():.2f} MW')

# 3. carrier별 발전기
print(f'\n3. Carrier별 발전기:')
for carrier in real_gens['carrier'].unique():
    gens = real_gens[real_gens['carrier']==carrier]
    print(f'   - {carrier}: {len(gens)}개, {gens["p_nom"].sum():.2f} MW')

# 4. 주요 발전기 (Nuclear, Coal, LNG)
print(f'\n4. 주요 발전기:')
for carrier in ['Nuclear', 'Coal', 'LNG']:
    gens = real_gens[real_gens['carrier']==carrier]
    if len(gens) > 0:
        print(f'   - {carrier}: {len(gens)}개, {gens["p_nom"].sum():.2f} MW')
        print(f'     p_min_pu 평균: {gens["p_min_pu"].mean():.4f}')

# 5. 구간 1 발전 가능량 추정
print(f'\n5. 구간 1 (732시간) 발전 가능량 추정:')

# Nuclear (p_min_pu=0.8로 가정)
nuclear_gens = real_gens[real_gens['carrier']=='Nuclear']
if len(nuclear_gens) > 0:
    nuclear_cap = nuclear_gens["p_nom"].sum()
    nuclear_gen = nuclear_cap * 0.8 * 732
    print(f'   - Nuclear (최소 80%): {nuclear_gen:.2f} MWh = {nuclear_gen/1e6:.2f} GWh')

# 재생에너지
re_gens = real_gens[real_gens['carrier'].isin(['solar', 'wind'])]
re_cap = re_gens["p_nom"].sum()
print(f'   - 재생에너지 (실제): 3.73 GWh')

# 총 필요 발전량
print(f'   - 총 수요: 71.69 GWh')
print(f'   - 부족분: {71.69 - 3.73:.2f} GWh')

# 6. 원자력 외 다른 발전기 용량
other_gens = real_gens[~real_gens['carrier'].isin(['Nuclear', 'solar', 'wind'])]
print(f'\n6. 원자력/재생에너지 외 발전기:')
print(f'   - 총 용량: {other_gens["p_nom"].sum():.2f} MW')
print(f'   - Carrier: {other_gens["carrier"].unique().tolist()}')

print('\n' + '='*70)
