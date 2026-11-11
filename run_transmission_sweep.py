"""
송전망 확장 연속 분석 (기준: 지역 상한 20% + 1인당 잠재량 비용 차등 α=0.25)

interface/통합 입력 파일은 수정하지 않고 메모리 상에서만 조건을 바꿔 연속 실행한다.

need_cost{x}  : 모든 AC 선로·HVDC 확장 허용(하한 = 기존 용량), 건설비 × x → 최적화가 필요한 구간만 증설
uniform_{p}   : 모든 AC 선로·HVDC 용량을 p% 일괄 증설(고정). 증설 비용은 같은 단가로 사후 산정

건설비 가정 (천원):
  AC 345kV  : 2,000 /MW·km (선로 길이 최소 30 km 적용)
  HVDC      : 3,000 /MW·km + 변환소 300,000 /MW
  연간화    : 40년, 할인율 4.5%, 운영유지비 1%/년

사용법:
  python run_transmission_sweep.py [--only 케이스 ...]
"""
import argparse
import os
import time
from datetime import datetime

import pandas as pd

import PyPSA_GUI as m
from run_ess_cost_sweep import summarize
from run_regional_cost_sweep import apply_potential_cap, apply_regional_cost, regional_capacity

AC_COST_PER_MW_KM = 2000.0
AC_MIN_LENGTH_KM = 30.0
HVDC_COST_PER_MW_KM = 3000.0
HVDC_CONVERTER_PER_MW = 300000.0
HVDC_LENGTH_KM = {'GWD_HVDC': 230.0, 'JBD_HVDC': 250.0, 'JND_HVDC': 400.0}
LIFETIME, RATE, FOM = 40, 0.045, 0.01
MAX_EXPANSION_FACTOR = 3.0  # 무한 상한은 Barrier 수렴 실패를 일으킴
ANNUITY = RATE / (1 - (1 + RATE) ** -LIFETIME) + FOM


def hvdc_links(n):
    return [l for l in HVDC_LENGTH_KM if l in n.links.index]


def annual_unit_costs(n):
    ac = n.lines.length.clip(lower=AC_MIN_LENGTH_KM) * AC_COST_PER_MW_KM * ANNUITY
    dc = pd.Series({l: (HVDC_LENGTH_KM[l] * HVDC_COST_PER_MW_KM + HVDC_CONVERTER_PER_MW) * ANNUITY
                    for l in hvdc_links(n)})
    return ac, dc


def apply_need_based(n, cost_mult):
    """AC 선로를 직접 확장 변수로 두면 KVL과 병렬 선로 때문에 Barrier가 수렴하지 못하므로,
    구간별 양방향 증설 링크(무손실, 조류 제어 가능 근사)를 추가하고 HVDC는 그대로 확장을 허용한다."""
    ac, dc = annual_unit_costs(n)
    L = n.lines
    if 'AC_expansion' not in n.carriers.index:
        n.add('Carrier', 'AC_expansion')
    for (b0, b1), idx in L.groupby([L.bus0, L.bus1]).groups.items():
        name = f"{b0.replace('_EL', '')}-{b1.replace('_EL', '')}_EXP"
        n.add('Link', name, bus0=b0, bus1=b1, carrier='AC_expansion', p_nom=0.0, p_nom_extendable=True,
              p_nom_max=L.loc[idx, 's_nom'].sum() * MAX_EXPANSION_FACTOR, p_min_pu=-1.0, efficiency=1.0,
              capital_cost=ac[idx].max() * cost_mult)
    hv = hvdc_links(n)
    n.links.loc[hv, 'p_nom_extendable'] = True
    n.links.loc[hv, 'p_nom_min'] = n.links.loc[hv, 'p_nom']
    n.links.loc[hv, 'p_nom_max'] = n.links.loc[hv, 'p_nom'] * MAX_EXPANSION_FACTOR
    n.links.loc[hv, 'capital_cost'] = dc * cost_mult
    print(f"[송전 확장] 필요 구간 확장 허용, 건설비 × {cost_mult} "
          f"(AC 평균 {ac.mean() * cost_mult:,.0f}, HVDC 평균 {dc.mean() * cost_mult:,.0f} 천원/MW·년)", flush=True)


def apply_uniform(n, pct):
    f = 1 + pct / 100
    n.lines['s_nom'] *= f
    hv = hvdc_links(n)
    n.links.loc[hv, 'p_nom'] *= f
    print(f"[송전 확장] 전 구간 일괄 +{pct}%", flush=True)


def corridor_table(n, base_lines, base_hvdc, label):
    ac, dc = annual_unit_costs(n)
    L = n.lines
    cap = L.s_nom_opt if L.s_nom_extendable.any() else L.s_nom
    add = cap - base_lines
    w = n.snapshot_weightings.generators
    p0 = n.lines_t.p0
    cong = (p0.abs().div(cap * L.s_max_pu)).ge(0.99).mul(w, axis=0).sum()
    rows = []
    for (b0, b1), idx in L.groupby([L.bus0, L.bus1]).groups.items():
        corridor = f"{b0.replace('_EL', '')}-{b1.replace('_EL', '')}"
        exp = f"{corridor}_EXP"
        if exp in n.links.index:
            added = float(n.links.at[exp, 'p_nom_opt'])
            cost = added * float(n.links.at[exp, 'capital_cost']) / 1e6
        else:
            added = float(add[idx].sum())
            cost = float((add[idx] * ac[idx]).sum()) / 1e6
        rows.append({'case': label, 'corridor': corridor, 'type': 'AC',
                     'existing_MW': base_lines[idx].sum(), 'final_MW': base_lines[idx].sum() + added,
                     'added_MW': added, 'annual_cost_bn': cost, 'congested_h': float(cong[idx].max())})
    for l in hvdc_links(n):
        c = n.links.at[l, 'p_nom_opt'] if n.links.at[l, 'p_nom_extendable'] else n.links.at[l, 'p_nom']
        flow = n.links_t.p0[l]
        rows.append({'case': label, 'corridor': f"{n.links.at[l, 'bus0'].replace('_EL', '')}-"
                                                f"{n.links.at[l, 'bus1'].replace('_EL', '')}", 'type': 'HVDC',
                     'existing_MW': base_hvdc[l], 'final_MW': c, 'added_MW': c - base_hvdc[l],
                     'annual_cost_bn': (c - base_hvdc[l]) * dc[l] / 1e6,
                     'congested_h': float((flow.abs() / c).ge(0.99).mul(w).sum())})
    return rows


def run_case(input_data, label, out_dir, kind, value):
    print(f"\n{'=' * 80}\n[연속 분석] {label}\n{'=' * 80}", flush=True)
    t0 = time.time()
    n = m.create_network(input_data)
    apply_potential_cap(n, 0.2)
    apply_regional_cost(n, 0.25, 'per_capita')
    base_lines = n.lines['s_nom'].copy()
    base_hvdc = n.links.loc[hvdc_links(n), 'p_nom'].copy()
    if kind == 'need':
        apply_need_based(n, value)
    elif kind == 'uniform':
        apply_uniform(n, value)
    ok = bool(m.optimize_network(n))
    row = summarize(n, label, 1.0, ok, time.time() - t0)
    row.update({'kind': kind, 'value': value})
    corr = []
    if ok:
        corr = corridor_table(n, base_lines, base_hvdc, label)
        c = pd.DataFrame(corr)
        row['grid_added_GW'] = c.added_MW.sum() / 1e3
        row['grid_annual_cost_bn'] = c.annual_cost_bn.sum()
        row['objective_incl_grid'] = row['objective'] + (c.annual_cost_bn.sum() * 1e6 if kind == 'uniform' else 0.0)
        case_dir = os.path.join(out_dir, label)
        os.makedirs(case_dir, exist_ok=True)
        n.export_to_netcdf(os.path.join(case_dir, f'{label}.nc'))
        reg = regional_capacity(n, label)
    else:
        reg = []
    print(f"[연속 분석] {label} 완료 ({row['elapsed_min']}분) PV={row.get('PV_GW', 0):.1f}GW "
          f"WT={row.get('WT_GW', 0):.1f}GW 출력제한={row.get('curtail_%', 0):.2f}% "
          f"송전증설={row.get('grid_added_GW', 0):.1f}GW", flush=True)
    return row, corr, reg


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--only', nargs='+')
    args = ap.parse_args()

    cases = [('ref_nogrid', 'none', 0)]
    cases += [(f"uniform_{p:02d}", 'uniform', p) for p in (10, 20, 30)]
    cases += [(f"need_cost{x:.1f}", 'need', x) for x in (1.0, 0.5, 2.0)]
    if args.only:
        cases = [c for c in cases if c[0] in args.only]

    out_dir = os.path.join('results', f"transmission_sweep_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    os.makedirs(out_dir, exist_ok=True)

    m.ensure_integrated_input()
    input_data = m.read_input_data(m.INPUT_FILE)
    input_data = m.standardize_bus_names_in_input(input_data)
    if 'generators' in input_data and not input_data['generators'].empty:
        input_data['generators']['committable'] = False

    rows, corrs, regs = [], [], []
    for label, kind, value in cases:
        try:
            row, corr, reg = run_case(input_data, label, out_dir, kind, value)
            rows.append(row); corrs += corr; regs += reg
        except Exception as e:
            print(f"[연속 분석] {label} 오류: {e}", flush=True)
            rows.append({'case': label, 'ok': False, 'error': str(e), 'kind': kind, 'value': value})
        pd.DataFrame(rows).to_csv(os.path.join(out_dir, 'sweep_summary.csv'), index=False, encoding='utf-8-sig')
        pd.DataFrame(corrs).to_csv(os.path.join(out_dir, 'corridor_expansion.csv'), index=False, encoding='utf-8-sig')
        pd.DataFrame(regs).to_csv(os.path.join(out_dir, 'regional_capacity.csv'), index=False, encoding='utf-8-sig')

    print(f"\n요약 파일: {out_dir}")


if __name__ == '__main__':
    main()
