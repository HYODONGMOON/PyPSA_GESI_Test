"""
석탄발전기 비발전 원인 진단 스크립트
- read_input_data → _replace_coal_generators_from_interface → standardize_bus_names 
  전 과정을 추적해서 coal 발전기 최종 파라미터를 확인
"""
import sys, os, pandas as pd
# src를 나중에 추가해 root의 PyPSA_GUI.py가 먼저 import되도록 함
sys.path.insert(0, os.path.abspath('.'))
# src는 나중에 (process_regional_excel.py 등 개별 모듈만 사용)
_src = os.path.abspath('src')
if _src not in sys.path:
    sys.path.append(_src)

LOG = []
def log(msg):
    print(msg)
    LOG.append(msg)

log("="*60)
log("석탄발전기 진단 시작")
log("="*60)

# ── 1) read_input_data 시뮬레이션 ───────────────────────────────
import pandas as pd
xls = pd.ExcelFile('integrated_input_data.xlsx')
input_data = {k: pd.read_excel('integrated_input_data.xlsx', sheet_name=k) for k in xls.sheet_names}

gen_df = input_data['generators']
coal_before = gen_df[gen_df['name'].astype(str).str.endswith('_Coal')]
log(f"\n[1] 로딩 전 집계 석탄발전기 ({len(coal_before)}개):")
for _, r in coal_before.iterrows():
    log(f"  {r['name']}: bus={r['bus']}, carrier={r['carrier']}, p_nom={r['p_nom']}, mc={r['marginal_cost']}")

# ── 2) _replace_coal_generators_from_interface (year=None 첫 호출) ──
from PyPSA_GUI import _replace_coal_generators_from_interface
interface_path = os.path.abspath('interface.xlsx')

log("\n[2] _replace_coal_generators_from_interface (year=None):")
try:
    input_data2 = _replace_coal_generators_from_interface(input_data.copy(), interface_path, year=None)
    gen2 = input_data2['generators']
    coal2 = gen2[gen2['carrier'].astype(str).str.lower() == 'coal']
    log(f"  → coal carrier 발전기: {len(coal2)}개")
    coal_endswith = gen2[gen2['name'].astype(str).str.endswith('_Coal')]
    log(f"  → *_Coal 집계발전기 잔존: {len(coal_endswith)}개")
    if not coal2.empty:
        log(coal2[['name','bus','carrier','p_nom','efficiency','marginal_cost']].to_string())
    else:
        log("  ★ coal carrier 발전기 없음!")
except Exception as e:
    log(f"  ★ 오류: {e}")
    import traceback; traceback.print_exc()
    input_data2 = input_data

# ── 3) _replace_coal_generators_from_interface (year=2030 재호출) ──
log("\n[3] _replace_coal_generators_from_interface (year=2030):")
try:
    input_data3 = _replace_coal_generators_from_interface(input_data2.copy(), interface_path, year=2030)
    gen3 = input_data3['generators']
    coal3 = gen3[gen3['carrier'].astype(str).str.lower() == 'coal']
    log(f"  → coal carrier 발전기: {len(coal3)}개")
    coal_endswith3 = gen3[gen3['name'].astype(str).str.endswith('_Coal')]
    log(f"  → *_Coal 집계발전기 잔존: {len(coal_endswith3)}개")
    if not coal3.empty:
        log(coal3[['name','bus','carrier','p_nom','efficiency','marginal_cost']].to_string())
    else:
        log("  ★ coal carrier 발전기 없음!")
except Exception as e:
    log(f"  ★ 오류: {e}")
    import traceback; traceback.print_exc()
    input_data3 = input_data2

# ── 4) standardize_bus_names_in_input ───────────────────────────
log("\n[4] standardize_bus_names_in_input 후:")
try:
    from PyPSA_GUI import standardize_bus_names_in_input
    input_data4 = standardize_bus_names_in_input(input_data3.copy())
    gen4 = input_data4['generators']
    coal4 = gen4[gen4['carrier'].astype(str).str.lower() == 'coal']
    log(f"  → coal 발전기 {len(coal4)}개:")
    if not coal4.empty:
        log(coal4[['name','bus','carrier','p_nom','efficiency','marginal_cost']].to_string())
    buses4 = input_data4['buses']['name'].dropna().tolist()
    log(f"  → EL 버스: {[b for b in buses4 if str(b).endswith('_EL')]}")
except Exception as e:
    log(f"  ★ 오류: {e}")

# ── 5) co2_emissions 계산 ──────────────────────────────────────
log("\n[5] CO2 배출량 계산 (석탄 vs LNG):")
log(f"  석탄 carrier co2_emissions = 0.73 * 0.47 = {0.73*0.47:.4f} tCO2/MWh_fuel")
log(f"  가스 carrier co2_emissions = 0.33 * 0.53 = {0.33*0.53:.4f} tCO2/MWh_fuel")
if not coal4.empty:
    for _, r in coal4.head(3).iterrows():
        eff = r['efficiency'] if pd.notna(r['efficiency']) else 0.47
        co2_per_mwhe = (0.73*0.47) / eff
        log(f"  {r['name']} (eff={eff:.3f}): CO2 = {co2_per_mwhe:.3f} tCO2/MWh_e, MC = {r['marginal_cost']}")
log(f"  LNG (eff=0.5): CO2 = {(0.33*0.53)/0.5:.3f} tCO2/MWh_e, MC = 226.02")
log(f"  → CO2 제약 mu=0 (비구속)이면 석탄(114.31) < LNG(226.02) → 석탄 우선 사용되어야 함!")

# ── 결과 저장 ──────────────────────────────────────────────────
with open('coal_diag_result.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(LOG))
log("\n진단 결과가 coal_diag_result.txt 에 저장되었습니다.")
