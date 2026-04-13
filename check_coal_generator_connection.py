import pandas as pd
import pypsa
import sys
import os

print("=" * 80)
print("1. integrated_input_data.xlsx에서 Coal 발전기 확인")
print("=" * 80)

# integrated_input_data.xlsx 읽기
df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
df_buses = pd.read_excel('integrated_input_data.xlsx', sheet_name='buses')

# Coal 발전기 필터링
coal_gens = df_gens[df_gens['carrier'] == 'coal'].copy()
print(f"\nCoal 발전기 수: {len(coal_gens)}")
print(f"Coal 발전기 총 용량: {coal_gens['p_nom'].sum():.1f} MW")

# 지역별 분포
print("\n지역별 Coal 발전기:")
for region in ['BSN', 'CBD', 'CND', 'DGU', 'DJN', 'GBD', 'GGD', 'GND', 'GWD', 'GWJ', 'ICN', 'JBD', 'JJD', 'JND', 'SEL', 'SJN', 'USN']:
    region_coal = coal_gens[coal_gens['name'].str.startswith(region + '_')]
    if not region_coal.empty:
        print(f"  {region}: {len(region_coal)}개, {region_coal['p_nom'].sum():.1f} MW")

print("\n" + "=" * 80)
print("2. Coal 발전기의 bus 연결 확인")
print("=" * 80)

# 각 Coal 발전기의 bus 확인
print("\nCoal 발전기가 연결된 버스:")
coal_buses = coal_gens['bus'].unique()
print(f"고유 버스 수: {len(coal_buses)}")
for bus in sorted(coal_buses):
    gen_count = len(coal_gens[coal_gens['bus'] == bus])
    print(f"  {bus}: {gen_count}개 발전기")

# 버스 존재 여부 확인
print("\n" + "=" * 80)
print("3. 버스 존재 여부 확인")
print("=" * 80)

available_buses = set(df_buses['name'].astype(str))
print(f"\n사용 가능한 버스 수: {len(available_buses)}")

missing_buses = []
for bus in coal_buses:
    if bus not in available_buses:
        missing_buses.append(bus)

if missing_buses:
    print(f"\n[WARNING] {len(missing_buses)}개의 버스가 buses 시트에 없습니다!")
    for bus in missing_buses:
        gen_count = len(coal_gens[coal_gens['bus'] == bus])
        print(f"  {bus}: {gen_count}개 발전기가 연결됨")
else:
    print("\n[OK] 모든 Coal 발전기의 버스가 buses 시트에 존재합니다.")

# 지역별 전력 버스 확인
print("\n" + "=" * 80)
print("4. 지역별 전력 버스 확인")
print("=" * 80)

# carrier가 'AC' 또는 'electricity'인 버스 찾기
el_buses = df_buses[df_buses['carrier'].isin(['AC', 'electricity'])].copy()
print(f"\n전력 버스 수: {len(el_buses)}")

# 지역별로 분류
print("\n지역별 전력 버스:")
for region in ['BSN', 'CBD', 'CND', 'DGU', 'DJN', 'GBD', 'GGD', 'GND', 'GWD', 'GWJ', 'ICN', 'JBD', 'JJD', 'JND', 'SEL', 'SJN', 'USN']:
    region_buses = el_buses[el_buses['name'].str.startswith(region + '_')]
    if not region_buses.empty:
        print(f"  {region}:")
        for bus_name in region_buses['name']:
            print(f"    - {bus_name}")

print("\n" + "=" * 80)
print("5. PyPSA Network 생성하여 실제 연결 확인")
print("=" * 80)

try:
    # PyPSA_GUI 모듈 임포트
    sys.path.insert(0, os.path.dirname(__file__))
    from PyPSA_GUI import read_input_data, create_network
    
    # 데이터 읽기
    print("\n데이터 로딩 중...")
    input_data = read_input_data('integrated_input_data.xlsx')
    
    # 네트워크 생성
    print("네트워크 생성 중...")
    network = create_network(input_data)
    
    # Coal 발전기 확인
    coal_gens_in_network = network.generators[network.generators.carrier == 'coal']
    print(f"\n[OK] Network에 추가된 Coal 발전기 수: {len(coal_gens_in_network)}")
    print(f"     총 용량: {coal_gens_in_network.p_nom.sum():.1f} MW")
    
    # 버스 연결 확인
    print("\nNetwork의 Coal 발전기 버스 연결:")
    coal_buses_in_network = coal_gens_in_network['bus'].unique()
    for bus in sorted(coal_buses_in_network):
        gen_count = len(coal_gens_in_network[coal_gens_in_network['bus'] == bus])
        # 해당 버스가 network에 있는지 확인
        if bus in network.buses.index:
            bus_carrier = network.buses.loc[bus, 'carrier']
            print(f"  {bus} (carrier: {bus_carrier}): {gen_count}개 발전기")
        else:
            print(f"  [WARNING] {bus}: 버스가 network에 없음! ({gen_count}개 발전기)")
    
    # 샘플 발전기 상세 정보
    print("\nCoal 발전기 샘플 (상위 10개):")
    print(coal_gens_in_network[['bus', 'p_nom', 'carrier', 'efficiency']].head(10))
    
    # 버스 통계
    print("\n" + "=" * 80)
    print("6. Network 버스 통계")
    print("=" * 80)
    print(f"\n총 버스 수: {len(network.buses)}")
    print(f"총 발전기 수: {len(network.generators)}")
    print(f"총 부하 수: {len(network.loads)}")
    
    # Carrier별 버스 분포
    print("\nCarrier별 버스 분포:")
    print(network.buses['carrier'].value_counts())
    
except Exception as e:
    print(f"\n[ERROR] 오류 발생: {str(e)}")
    import traceback
    traceback.print_exc()
