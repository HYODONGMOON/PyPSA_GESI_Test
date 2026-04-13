"""
Unbounded 문제의 근본 원인을 진단하는 스크립트
"""
import pandas as pd
import numpy as np
import sys
import io

# 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

print("\n" + "="*70)
print("Unbounded 문제 원인 진단")
print("="*70 + "\n")

# integrated_input_data.xlsx 파일 로드
file_path = r"C:\Users\Hyodong.Moon\Desktop\HDMOON\python workplace\PyPSA_GESI_Test\integrated_input_data.xlsx"

print(f"파일 로드 중: {file_path}\n")

# 1. Generators 시트 확인
print("\n[1] Generators 시트 확인")
print("-" * 70)
generators = pd.read_excel(file_path, sheet_name='generators')
print(f"총 발전기 수: {len(generators)}")

# 1-1. Marginal cost가 음수이거나 0인 발전기 (extendable 제외)
print("\n[1-1] Marginal cost가 0 이하인 발전기:")
zero_marginal = generators[generators['marginal_cost'] <= 0]
if len(zero_marginal) > 0:
    print(f"  [경고] {len(zero_marginal)}개 발견!")
    print(zero_marginal[['name', 'carrier', 'marginal_cost', 'p_nom', 'p_nom_extendable']].head(20))
else:
    print("  [OK] 문제 없음")

# 1-2. Extendable이면서 capital_cost가 0 이하인 발전기
print("\n[1-2] Extendable이면서 capital_cost가 0 이하인 발전기:")
zero_capital = generators[(generators['p_nom_extendable'] == True) & (generators['capital_cost'] <= 0)]
if len(zero_capital) > 0:
    print(f"  [경고] {len(zero_capital)}개 발견!")
    print(zero_capital[['name', 'carrier', 'marginal_cost', 'capital_cost', 'p_nom_extendable']].head(20))
else:
    print("  [OK] 문제 없음")

# 1-3. Efficiency가 1보다 큰 발전기
print("\n[1-3] Efficiency가 1보다 큰 발전기:")
high_eff_gen = generators[generators['efficiency'] > 1.0]
if len(high_eff_gen) > 0:
    print(f"  [경고] {len(high_eff_gen)}개 발견!")
    print(high_eff_gen[['name', 'carrier', 'efficiency', 'marginal_cost']].head(20))
else:
    print("  [OK] 문제 없음")

# 2. Links 시트 확인
print("\n[2] Links 시트 확인")
print("-" * 70)
links = pd.read_excel(file_path, sheet_name='links')
print(f"총 링크 수: {len(links)}")

# 2-1. Marginal cost가 음수인 링크
print("\n[2-1] Marginal cost가 음수인 링크:")
neg_marginal_link = links[links['marginal_cost'] < 0]
if len(neg_marginal_link) > 0:
    print(f"  [경고] {len(neg_marginal_link)}개 발견!")
    print(neg_marginal_link[['name', 'bus0', 'bus1', 'marginal_cost']].head(20))
else:
    print("  [OK] 문제 없음")

# 2-2. Extendable이면서 capital_cost가 0 이하인 링크
print("\n[2-2] Extendable이면서 capital_cost가 0 이하인 링크:")
zero_capital_link = links[(links['p_nom_extendable'] == True) & (links['capital_cost'] <= 0)]
if len(zero_capital_link) > 0:
    print(f"  [경고] {len(zero_capital_link)}개 발견!")
    print(zero_capital_link[['name', 'bus0', 'bus1', 'capital_cost', 'p_nom_extendable']].head(20))
else:
    print("  [OK] 문제 없음")

# 2-3. Efficiency가 1보다 큰 링크 (에너지 생성)
print("\n[2-3] Efficiency가 1보다 큰 링크:")
for eff_col in ['efficiency0', 'efficiency1', 'efficiency2', 'efficiency3']:
    if eff_col in links.columns:
        high_eff_links = links[links[eff_col] > 1.0]
        if len(high_eff_links) > 0:
            print(f"  [경고] {eff_col}에서 {len(high_eff_links)}개 발견!")
            print(high_eff_links[['name', 'bus0', 'bus1', eff_col]].head(10))

# 3. Stores 시트 확인
print("\n[3] Stores 시트 확인")
print("-" * 70)
stores = pd.read_excel(file_path, sheet_name='stores')
print(f"총 저장소 수: {len(stores)}")

# 3-1. Marginal cost가 음수인 저장소
print("\n[3-1] Marginal cost가 음수인 저장소:")
neg_marginal_store = stores[stores['marginal_cost'] < 0]
if len(neg_marginal_store) > 0:
    print(f"  [경고] {len(neg_marginal_store)}개 발견!")
    print(neg_marginal_store[['name', 'bus', 'marginal_cost']].head(20))
else:
    print("  [OK] 문제 없음")

# 3-2. Efficiency가 1보다 큰 저장소
print("\n[3-2] Efficiency (store/dispatch)가 1보다 큰 저장소:")
high_eff_store_in = stores[stores['efficiency_store'] > 1.0]
high_eff_store_out = stores[stores['efficiency_dispatch'] > 1.0]
if len(high_eff_store_in) > 0 or len(high_eff_store_out) > 0:
    print(f"  [경고] Store: {len(high_eff_store_in)}개, Dispatch: {len(high_eff_store_out)}개 발견!")
    if len(high_eff_store_in) > 0:
        print("  [efficiency_store > 1.0]")
        print(high_eff_store_in[['name', 'bus', 'efficiency_store', 'efficiency_dispatch']].head(10))
    if len(high_eff_store_out) > 0:
        print("  [efficiency_dispatch > 1.0]")
        print(high_eff_store_out[['name', 'bus', 'efficiency_store', 'efficiency_dispatch']].head(10))
else:
    print("  [OK] 문제 없음")

# 3-3. Extendable이면서 capital_cost가 0 이하인 저장소
print("\n[3-3] Extendable이면서 capital_cost가 0 이하인 저장소:")
zero_capital_store = stores[(stores['e_nom_extendable'] == True) & (stores['capital_cost'] <= 0)]
if len(zero_capital_store) > 0:
    print(f"  [경고] {len(zero_capital_store)}개 발견!")
    print(zero_capital_store[['name', 'bus', 'capital_cost', 'e_nom_extendable']].head(20))
else:
    print("  [OK] 문제 없음")

# 4. Lines 시트 확인
print("\n[4] Lines 시트 확인")
print("-" * 70)
lines = pd.read_excel(file_path, sheet_name='lines')
print(f"총 선로 수: {len(lines)}")

# 4-1. 리액턴스(x)가 0 이하인 선로
print("\n[4-1] 리액턴스(x)가 0 이하인 선로:")
zero_reactance = lines[lines['x'] <= 0]
if len(zero_reactance) > 0:
    print(f"  [경고] {len(zero_reactance)}개 발견!")
    print(zero_reactance[['name', 'bus0', 'bus1', 'x', 'r', 's_nom']].head(20))
else:
    print("  [OK] 문제 없음")

# 4-2. 저항(r)이 음수인 선로
print("\n[4-2] 저항(r)이 음수인 선로:")
neg_resistance = lines[lines['r'] < 0]
if len(neg_resistance) > 0:
    print(f"  [경고] {len(neg_resistance)}개 발견!")
    print(neg_resistance[['name', 'bus0', 'bus1', 'r', 'x', 's_nom']].head(20))
else:
    print("  [OK] 문제 없음")

# 4-3. s_nom이 0 이하인 선로
print("\n[4-3] s_nom이 0 이하인 선로:")
zero_snom = lines[lines['s_nom'] <= 0]
if len(zero_snom) > 0:
    print(f"  [경고] {len(zero_snom)}개 발견!")
    print(zero_snom[['name', 'bus0', 'bus1', 's_nom', 'r', 'x']].head(20))
else:
    print("  [OK] 문제 없음")

# 5. 요약
print("\n" + "="*70)
print("진단 요약")
print("="*70)
print("\n발견된 잠재적 문제:")
issues = []

if len(zero_marginal) > 0:
    issues.append(f"  • Marginal cost ≤ 0인 발전기: {len(zero_marginal)}개")
    
if len(zero_capital) > 0:
    issues.append(f"  • Extendable + capital_cost ≤ 0인 발전기: {len(zero_capital)}개")
    
if len(high_eff_gen) > 0:
    issues.append(f"  • Efficiency > 1.0인 발전기: {len(high_eff_gen)}개")
    
if len(neg_marginal_link) > 0:
    issues.append(f"  • Marginal cost < 0인 링크: {len(neg_marginal_link)}개")
    
if len(zero_capital_link) > 0:
    issues.append(f"  • Extendable + capital_cost ≤ 0인 링크: {len(zero_capital_link)}개")
    
if len(zero_capital_store) > 0:
    issues.append(f"  • Extendable + capital_cost ≤ 0인 저장소: {len(zero_capital_store)}개")
    
if len(zero_reactance) > 0:
    issues.append(f"  • 리액턴스(x) ≤ 0인 선로: {len(zero_reactance)}개")
    
if len(neg_resistance) > 0:
    issues.append(f"  • 저항(r) < 0인 선로: {len(neg_resistance)}개")
    
if len(zero_snom) > 0:
    issues.append(f"  • s_nom ≤ 0인 선로: {len(zero_snom)}개")

if issues:
    for issue in issues:
        print(issue)
else:
    print("  [OK] 명백한 문제가 발견되지 않았습니다.")

print("\n" + "="*70)
print("진단 완료!")
print("="*70 + "\n")
