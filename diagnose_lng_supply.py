"""
LNG 버스 공급원 진단
"""
import pandas as pd
import numpy as np
from PyPSA_GUI import read_input_data, create_network, standardize_bus_names_in_input

# 데이터 로드 및 네트워크 생성
input_data = read_input_data('integrated_input_data.xlsx')
input_data = standardize_bus_names_in_input(input_data)
network = create_network(input_data)

print('='*70)
print('LNG 버스 공급원 진단')
print('='*70)

# 1. LNG 버스 확인
lng_buses = [b for b in network.buses.index if 'LNG' in b]
print(f'\n1. LNG 버스: {len(lng_buses)}개')
for bus in lng_buses[:5]:
    print(f'   - {bus}')

# 2. LNG 버스에 연결된 발전기 (공급원)
print(f'\n2. LNG 버스에 연결된 발전기:')
lng_gens = network.generators[network.generators['bus'].isin(lng_buses)]
if len(lng_gens) > 0:
    print(f'   - 총 {len(lng_gens)}개')
    for idx, gen in lng_gens.iterrows():
        print(f'   - {idx}: bus={gen["bus"]}, carrier={gen["carrier"]}, p_nom={gen["p_nom"]:.2f} MW')
else:
    print('   ❌ LNG 버스에 발전기 없음!')

# 3. LNG 버스에 연결된 stores (저장소)
print(f'\n3. LNG 버스에 연결된 stores:')
lng_stores = network.stores[network.stores['bus'].isin(lng_buses)]
if len(lng_stores) > 0:
    print(f'   - 총 {len(lng_stores)}개')
    for idx, store in lng_stores.iterrows():
        print(f'   - {idx}: bus={store["bus"]}, e_nom={store["e_nom"]:.2f} MWh')
else:
    print('   ❌ LNG 버스에 stores 없음!')

# 4. LNG 버스를 bus1로 가지는 Links (LNG로 들어오는 흐름)
print(f'\n4. LNG 버스로 들어오는 Links (bus1 = LNG):')
incoming_links = network.links[network.links['bus1'].isin(lng_buses)]
if len(incoming_links) > 0:
    print(f'   - 총 {len(incoming_links)}개')
    for idx, link in incoming_links.iterrows():
        print(f'   - {idx}: {link["bus0"]} → {link["bus1"]}, p_nom={link["p_nom"]:.2f} MW')
else:
    print('   ❌ LNG 버스로 들어오는 Links 없음!')

# 5. CHP Links (LNG를 소비하는 Links)
print(f'\n5. CHP Links (bus0 = LNG, LNG를 소비):')
chp_links = network.links[network.links['bus0'].isin(lng_buses)]
if len(chp_links) > 0:
    print(f'   - 총 {len(chp_links)}개')
    total_chp_capacity = chp_links['p_nom'].sum()
    print(f'   - 총 CHP 용량: {total_chp_capacity:.2f} MW')
    for idx, link in chp_links[:5].iterrows():
        bus2_info = f", bus2={link['bus2']}" if pd.notna(link.get('bus2')) else ""
        print(f'   - {idx}: {link["bus0"]} → {link["bus1"]}{bus2_info}, p_nom={link["p_nom"]:.2f} MW')
else:
    print('   ❌ CHP Links 없음!')

# 6. 문제 진단
print(f'\n' + '='*70)
print('💡 진단 결과:')
print('='*70)

if len(lng_gens) == 0 and len(lng_stores) == 0 and len(incoming_links) == 0:
    print('❌ 심각한 문제: LNG 버스에 공급원이 전혀 없습니다!')
    print('   → CHP가 LNG를 소비하려고 하지만, LNG 공급이 불가능합니다.')
    print('   → 이것이 infeasible의 근본 원인입니다!')
    print('\n해결 방법:')
    print('   1. LNG 버스에 무한 용량 발전기 추가 (p_nom_extendable=True)')
    print('   2. 또는 LNG stores 추가 (충분한 e_nom)')
    print('   3. 또는 외부에서 LNG로 들어오는 Links 추가')
else:
    if len(lng_gens) > 0:
        total_gen_cap = lng_gens['p_nom'].sum()
        print(f'✓ LNG 발전기 용량: {total_gen_cap:.2f} MW')
        if total_gen_cap < total_chp_capacity:
            print(f'⚠️ LNG 공급 부족: CHP 용량 {total_chp_capacity:.2f} MW > LNG 공급 {total_gen_cap:.2f} MW')
    if len(lng_stores) > 0:
        total_store_cap = lng_stores['e_nom'].sum()
        print(f'✓ LNG stores 용량: {total_store_cap:.2f} MWh')
    if len(incoming_links) > 0:
        print(f'✓ LNG로 들어오는 Links: {len(incoming_links)}개')

print('\n' + '='*70)
