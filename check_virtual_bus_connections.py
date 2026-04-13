"""
가상버스 연결 송전선 용량 확인
"""
import pandas as pd
import sys
import io

# 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

file_path = r"C:\Users\Hyodong.Moon\Desktop\HDMOON\python workplace\PyPSA_GESI_Test\integrated_input_data.xlsx"

print("\n" + "="*70)
print("가상버스 연결 송전선 분석")
print("="*70 + "\n")

# 버스 로드
buses = pd.read_excel(file_path, sheet_name='buses')
lines = pd.read_excel(file_path, sheet_name='lines')

# 지역 대표 버스 (가상버스) 찾기
virtual_buses = buses[buses['name'].str.match(r'^[A-Z]{3}_EL$', na=False)]['name'].tolist()
print(f"지역 대표 버스 (가상버스): {len(virtual_buses)}개")
print(f"예시: {virtual_buses[:5]}\n")

# 실제 송전망 버스 찾기
physical_buses = buses[buses['name'].str.contains('bus_', na=False)]['name'].tolist()
print(f"실제 송전망 버스: {len(physical_buses)}개")
print(f"예시: {physical_buses[:5]}\n")

# 가상버스와 연결된 송전선 찾기
print("="*70)
print("가상버스 연결 송전선 분석")
print("="*70 + "\n")

virtual_connections = []

for vbus in virtual_buses:
    # 해당 가상버스에 연결된 송전선 찾기
    connected_lines = lines[
        (lines['bus0'] == vbus) | (lines['bus1'] == vbus)
    ]
    
    if len(connected_lines) > 0:
        print(f"\n[{vbus}] 연결 송전선: {len(connected_lines)}개")
        
        # 송전선 용량 통계
        capacities = connected_lines['s_nom'].values
        print(f"  용량 통계:")
        print(f"    - 최소: {capacities.min():.2f} MVA")
        print(f"    - 최대: {capacities.max():.2f} MVA")
        print(f"    - 평균: {capacities.mean():.2f} MVA")
        print(f"    - 총합: {capacities.sum():.2f} MVA")
        
        # 샘플 출력 (최대 5개)
        print(f"\n  샘플 송전선:")
        for idx, row in connected_lines.head(5).iterrows():
            print(f"    {row['name']}: {row['bus0']} ↔ {row['bus1']}, {row['s_nom']:.2f} MVA")
        
        virtual_connections.extend(connected_lines.to_dict('records'))

# 전체 통계
print("\n" + "="*70)
print("전체 가상버스 연결 통계")
print("="*70)

if virtual_connections:
    all_capacities = [c['s_nom'] for c in virtual_connections]
    print(f"\n총 가상버스 연결 송전선: {len(virtual_connections)}개")
    print(f"용량 통계:")
    print(f"  - 최소: {min(all_capacities):.2f} MVA")
    print(f"  - 최대: {max(all_capacities):.2f} MVA")
    print(f"  - 평균: {sum(all_capacities)/len(all_capacities):.2f} MVA")
    print(f"  - 중앙값: {sorted(all_capacities)[len(all_capacities)//2]:.2f} MVA")
else:
    print("\n[경고] 가상버스에 직접 연결된 송전선이 없습니다!")
    print("이것이 문제의 원인일 수 있습니다.")

# 지역별 수요 확인
print("\n" + "="*70)
print("지역별 수요 대비 송전 용량 확인")
print("="*70 + "\n")

loads = pd.read_excel(file_path, sheet_name='loads')

for vbus in virtual_buses[:10]:  # 처음 10개만
    region = vbus.replace('_EL', '')
    
    # 해당 지역 수요
    region_loads = loads[loads['bus'] == vbus]
    if len(region_loads) > 0:
        total_demand = region_loads['p_set'].sum()
        
        # 연결 용량
        connected_lines = lines[(lines['bus0'] == vbus) | (lines['bus1'] == vbus)]
        total_capacity = connected_lines['s_nom'].sum() if len(connected_lines) > 0 else 0
        
        ratio = total_capacity / total_demand if total_demand > 0 else 0
        
        status = "✓" if ratio >= 1.5 else "⚠" if ratio >= 1.0 else "✗"
        print(f"{status} {vbus}: 수요 {total_demand:.0f} MW, 송전용량 {total_capacity:.0f} MVA (비율: {ratio:.2f})")

print("\n" + "="*70)
print("분석 완료!")
print("="*70 + "\n")
