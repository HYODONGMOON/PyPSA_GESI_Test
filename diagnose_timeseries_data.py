"""
시계열 데이터와 네트워크 구조 진단
"""
import pandas as pd
import numpy as np
import sys
import io

# 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("\n" + "="*70)
print("시계열 데이터 및 네트워크 구조 진단")
print("="*70 + "\n")

file_path = r"C:\Users\Hyodong.Moon\Desktop\HDMOON\python workplace\PyPSA_GESI_Test\integrated_input_data.xlsx"

# 1. Renewable patterns 확인
print("[1] 재생에너지 패턴 확인")
print("-" * 70)
try:
    renewable_patterns = pd.read_excel(file_path, sheet_name='renewable_patterns')
    print(f"행 수: {len(renewable_patterns)}")
    print(f"컬럼 수: {len(renewable_patterns.columns)}")
    
    # 첫 번째 행이 헤더인지 확인
    if '재생에너지 발전 패턴 설정' in renewable_patterns.columns:
        print("[주의] 첫 번째 행이 헤더로 인식됨")
        # 실제 데이터 행 확인
        actual_data = renewable_patterns.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')
        print(f"\n통계:")
        print(f"  - 최소값: {actual_data.min().min()}")
        print(f"  - 최대값: {actual_data.max().max()}")
        print(f"  - 평균값: {actual_data.mean().mean()}")
        print(f"  - NaN 개수: {actual_data.isna().sum().sum()}")
        print(f"  - 0인 값 개수: {(actual_data == 0).sum().sum()}")
        
        # 모든 값이 0인 컬럼 확인
        zero_cols = actual_data.columns[(actual_data == 0).all()]
        if len(zero_cols) > 0:
            print(f"\n  [경고] 모든 값이 0인 컬럼: {len(zero_cols)}개")
            print(f"  컬럼: {list(zero_cols[:10])}")
except Exception as e:
    print(f"[오류] {str(e)}")

# 2. Load patterns 확인
print("\n[2] 부하 패턴 확인")
print("-" * 70)
try:
    load_patterns = pd.read_excel(file_path, sheet_name='load_patterns')
    print(f"행 수: {len(load_patterns)}")
    print(f"컬럼 수: {len(load_patterns.columns)}")
    
    if '부하 패턴 설정 (국가 기본값)' in load_patterns.columns:
        print("[주의] 첫 번째 행이 헤더로 인식됨")
        actual_data = load_patterns.iloc[:, 1:].apply(pd.to_numeric, errors='coerce')
        print(f"\n통계:")
        print(f"  - 최소값: {actual_data.min().min()}")
        print(f"  - 최대값: {actual_data.max().max()}")
        print(f"  - 평균값: {actual_data.mean().mean()}")
        print(f"  - NaN 개수: {actual_data.isna().sum().sum()}")
except Exception as e:
    print(f"[오류] {str(e)}")

# 3. Generators 확인
print("\n[3] 발전기 상세 확인")
print("-" * 70)
generators = pd.read_excel(file_path, sheet_name='generators')

# Committable 발전기 확인
committable_gens = generators[generators['committable'] == True]
print(f"Committable 발전기: {len(committable_gens)}개")
if len(committable_gens) > 0:
    print("\n주요 속성:")
    print(committable_gens[['name', 'carrier', 'p_nom', 'marginal_cost', 'ramp_limit_up', 'ramp_limit_down']].head(10))
    
    # Ramp limit이 너무 작은지 확인
    if 'ramp_limit_up' in committable_gens.columns:
        low_ramp = committable_gens[committable_gens['ramp_limit_up'] < 0.1]
        if len(low_ramp) > 0:
            print(f"\n  [경고] Ramp limit이 0.1 미만인 발전기: {len(low_ramp)}개")

# p_min_pu 확인
if 'p_min_pu' in generators.columns:
    high_pmin = generators[generators['p_min_pu'] > 0.5]
    if len(high_pmin) > 0:
        print(f"\n  [경고] p_min_pu가 0.5 초과인 발전기: {len(high_pmin)}개")
        print(high_pmin[['name', 'carrier', 'p_min_pu', 'committable']].head(10))

# 4. Constraints 확인
print("\n[4] 제약조건 확인")
print("-" * 70)
try:
    constraints = pd.read_excel(file_path, sheet_name='constraints')
    print(f"총 제약조건 수: {len(constraints)}")
    print("\n제약조건 목록:")
    print(constraints[['name', 'type', 'sense', 'constant', 'description']])
    
    # CO2 제약 상세 확인
    co2_constraints = constraints[constraints['name'].str.contains('CO2', na=False)]
    if len(co2_constraints) > 0:
        print(f"\n  CO2 제약:")
        for idx, row in co2_constraints.iterrows():
            print(f"    - {row['name']}: {row['sense']} {row['constant']}")
except Exception as e:
    print(f"[오류] {str(e)}")

# 5. Loads 확인
print("\n[5] 부하 확인")
print("-" * 70)
loads = pd.read_excel(file_path, sheet_name='loads')
print(f"총 부하 수: {len(loads)}")
print(f"부하별 연결 버스:")
print(loads[['name', 'bus', 'carrier', 'p_set']].head(20))

# p_set이 너무 큰지 확인
if loads['p_set'].dtype in [np.float64, np.int64]:
    print(f"\n부하 통계:")
    print(f"  - 최소값: {loads['p_set'].min()}")
    print(f"  - 최대값: {loads['p_set'].max()}")
    print(f"  - 평균값: {loads['p_set'].mean()}")
    print(f"  - 총합: {loads['p_set'].sum()}")

# 6. 버스 연결성 확인
print("\n[6] 버스 연결성 확인")
print("-" * 70)
buses = pd.read_excel(file_path, sheet_name='buses')
lines = pd.read_excel(file_path, sheet_name='lines')

print(f"총 버스 수: {len(buses)}")
print(f"총 선로 수: {len(lines)}")

# 전력 버스만 추출
el_buses = buses[buses['carrier'] == 'AC']
print(f"전력(AC) 버스 수: {len(el_buses)}")

# 각 전력 버스에 연결된 선로 수 확인
bus_connections = {}
for bus in el_buses['name']:
    connected_lines = len(lines[(lines['bus0'] == bus) | (lines['bus1'] == bus)])
    if connected_lines == 0:
        if bus not in bus_connections:
            bus_connections[bus] = connected_lines

isolated_buses = [b for b, c in bus_connections.items() if c == 0]
if isolated_buses:
    print(f"\n  [경고] 선로가 연결되지 않은 전력 버스: {len(isolated_buses)}개")
    print(f"  예시: {isolated_buses[:10]}")

print("\n" + "="*70)
print("진단 완료!")
print("="*70 + "\n")
