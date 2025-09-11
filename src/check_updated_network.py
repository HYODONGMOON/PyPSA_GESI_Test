import pandas as pd
import numpy as np

# 데이터 읽기
buses = pd.read_excel('integrated_input_data.xlsx', sheet_name='buses')
lines = pd.read_excel('integrated_input_data.xlsx', sheet_name='lines')
generators = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
loads = pd.read_excel('integrated_input_data.xlsx', sheet_name='loads')

print(f'버스 수: {len(buses)}, 라인 수: {len(lines)}, 발전기 수: {len(generators)}, 부하 수: {len(loads)}')

# 버스 이름 목록
bus_names = set(buses['name'])

# 각 지역에 발전기, 버스, 부하가 제대로 설정되어 있는지 확인
regions = sorted(list(set([name.split('_')[0] for name in bus_names if '_' in name])))
print(f"\n총 지역 수: {len(regions)}")
print(f"지역 목록: {regions}")

# 각 지역의 컴포넌트 수 확인
region_components = {}
for region in regions:
    region_buses = [bus for bus in bus_names if bus.startswith(region + '_')]
    region_generators = generators[generators['bus'].apply(lambda x: str(x).startswith(region + '_') if isinstance(x, str) else False)]
    region_loads = loads[loads['bus'].apply(lambda x: str(x).startswith(region + '_') if isinstance(x, str) else False)]
    
    region_components[region] = {
        'buses': len(region_buses),
        'generators': len(region_generators),
        'loads': len(region_loads),
    }

print("\n지역별 컴포넌트 수:")
for region, components in region_components.items():
    print(f"{region}: 버스 {components['buses']}개, 발전기 {components['generators']}개, 부하 {components['loads']}개")

# 지역간 연결 확인 (전력망)
region_connections = []
for _, line in lines.iterrows():
    bus0 = str(line['bus0'])
    bus1 = str(line['bus1'])
    
    # 버스 이름에서 지역 코드 추출 (예: BSN_EL에서 BSN)
    region0 = bus0.split('_')[0] if '_' in bus0 else bus0
    region1 = bus1.split('_')[0] if '_' in bus1 else bus1
    
    if region0 != region1:
        connection = f"{region0}-{region1}"
        region_connections.append(connection)

unique_connections = set(region_connections)
print(f"\n지역간 전력망 연결 수: {len(unique_connections)}")
print(f"지역간 연결 목록: {sorted(list(unique_connections))}")

# 단위 확인
print("\n===== 데이터 단위 확인 =====")

print("\n라인 데이터의 단위:")
print(f"s_nom (송전용량) 단위: {'MW 로 추정됨' if lines['s_nom'].max() > 100 else 'GW 로 추정됨'}")
print(f"v_nom (전압) 단위: {'kV 로 추정됨' if lines['v_nom'].max() > 100 else 'V 로 추정됨'}")
print(f"length (길이) 단위: {'km 로 추정됨' if lines['length'].max() > 10 else 'm 로 추정됨'}")
print(f"r (저항) 범위: {lines['r'].min()} ~ {lines['r'].max()}")
print(f"x (리액턴스) 범위: {lines['x'].min()} ~ {lines['x'].max()}")

print("\n발전기 데이터의 단위:")
print(f"p_nom (발전용량) 단위: {'MW 로 추정됨' if generators['p_nom'].max() > 100 else 'GW 로 추정됨'}")
if 'capital_cost' in generators.columns:
    not_nan_values = generators['capital_cost'].dropna()
    if len(not_nan_values) > 0:
        print(f"capital_cost (설비투자비) 범위: {not_nan_values.min()} ~ {not_nan_values.max()}")
    else:
        print("capital_cost (설비투자비): 모든 값이 NaN")

if 'marginal_cost' in generators.columns:
    not_nan_values = generators['marginal_cost'].dropna()
    if len(not_nan_values) > 0:
        print(f"marginal_cost (한계비용) 범위: {not_nan_values.min()} ~ {not_nan_values.max()}")
    else:
        print("marginal_cost (한계비용): 모든 값이 NaN")

print("\n부하 데이터의 단위:")
print(f"p_set (전력수요) 단위: {'MW 로 추정됨' if loads['p_set'].max() > 100 else 'GW 로 추정됨'}")

# 버스별 발전기와 부하 확인
el_buses = [bus for bus in bus_names if bus.endswith('_EL')]
print(f"\n전력 버스 (EL) 개수: {len(el_buses)}")

# 각 지역의 전력 버스에 발전기와 부하가 있는지 확인
missing_generators = []
missing_loads = []

for bus in el_buses:
    if len(generators[generators['bus'] == bus]) == 0:
        missing_generators.append(bus)
    if len(loads[loads['bus'] == bus]) == 0:
        missing_loads.append(bus)

if missing_generators:
    print(f"\n발전기가 없는 전력 버스: {missing_generators}")
else:
    print("\n모든 전력 버스에 발전기가 연결되어 있습니다.")

if missing_loads:
    print(f"\n부하가 없는 전력 버스: {missing_loads}")
else:
    print("\n모든 전력 버스에 부하가 연결되어 있습니다.")

# 라인 연결 확인
unconnected_buses = []
for bus in el_buses:
    outgoing = len(lines[lines['bus0'] == bus])
    incoming = len(lines[lines['bus1'] == bus])
    if outgoing + incoming == 0:
        unconnected_buses.append(bus)

if unconnected_buses:
    print(f"\n라인이 연결되지 않은 전력 버스: {unconnected_buses}")
else:
    print("\n모든 전력 버스가 라인으로 연결되어 있습니다.")

# 버스간 연결 시각화 (단순화된 방식)
connections = {}
for _, line in lines.iterrows():
    bus0 = line['bus0']
    bus1 = line['bus1']
    if bus0 not in connections:
        connections[bus0] = []
    if bus1 not in connections:
        connections[bus1] = []
    connections[bus0].append(bus1)
    connections[bus1].append(bus0)

# 각 버스가 다른 모든 버스로 갈 수 있는지 확인 (간단한 BFS)
def is_connected(start_bus, connections):
    visited = set()
    queue = [start_bus]
    while queue:
        bus = queue.pop(0)
        visited.add(bus)
        for neighbor in connections.get(bus, []):
            if neighbor not in visited:
                queue.append(neighbor)
    return visited

# 전력 버스들이 서로 연결되어 있는지 확인
el_buses_set = set(el_buses)
if el_buses:
    connected_buses = is_connected(el_buses[0], connections)
    disconnected = el_buses_set - connected_buses
    if disconnected:
        print(f"\n전력망에서 분리된 버스: {list(disconnected)}")
    else:
        print("\n모든 전력 버스가 서로 연결되어 있습니다 (전력망이 완전히 연결됨).")
else:
    print("\n전력 버스가 없습니다.")

print("\n===== 네트워크 분석 요약 =====")
print(f"총 버스 수: {len(buses)}")
print(f"총 지역 수: {len(regions)}")
print(f"지역간 연결 수: {len(unique_connections)}")

if not missing_generators and not missing_loads and not unconnected_buses and el_buses and not (el_buses_set - is_connected(el_buses[0], connections)):
    print("\n네트워크가 잘 연결되어 있고 문제가 발견되지 않았습니다. 분석을 진행해도 좋을 것 같습니다.")
else:
    print("\n네트워크에 일부 문제가 있을 수 있습니다. 위의 경고 메시지를 확인하고 필요한 경우 데이터를 수정하는 것이 좋겠습니다.") 