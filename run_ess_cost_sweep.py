"""
ESS 비용 민감도 연속 분석

interface/통합 입력 파일은 수정하지 않고, 메모리 상에서 ESS 비용(저장용량 *_ESS Store +
충방전 *_ESS_Ch/*_ESS_Disch Link의 capital_cost)에만 배율을 곱해 연속 실행한다.

사용법:
  python run_ess_cost_sweep.py                      # 기준(1.0) 실행 → 손익분기 배율 k* 추정 → k*, -10/-20/-30% 실행
  python run_ess_cost_sweep.py --multipliers 0.6 0.5 # 지정한 배율만 실행 (기준 실행 생략)
"""
import argparse
import copy
import json
import os
import sys
import time
from datetime import datetime

import numpy as np
import pandas as pd

import PyPSA_GUI as m

STEPS = (0.0, 0.10, 0.20, 0.30)
DURATIONS_H = (1, 2, 4, 6, 8)


def scale_ess_cost(input_data, mult):
    data = copy.deepcopy(input_data)
    st = data['stores']
    mask = st['name'].astype(str).str.endswith('_ESS')
    st.loc[mask, 'capital_cost'] = st.loc[mask, 'capital_cost'].astype(float) * mult
    lk = data['links']
    lmask = lk['name'].astype(str).str.contains(r'_ESS_(?:Ch|Disch)$', regex=True)
    lk.loc[lmask, 'capital_cost'] = lk.loc[lmask, 'capital_cost'].astype(float) * mult
    return data


def summarize(n, label, mult, ok, elapsed):
    row = {'case': label, 'ess_cost_mult': mult, 'ok': ok, 'elapsed_min': round(elapsed / 60, 1)}
    if not ok:
        return row
    g, P = n.generators, n.generators_t.p
    row['objective'] = float(n.objective)
    for car, key in (('solar', 'PV'), ('wind', 'WT')):
        idx = g.index[g.carrier == car]
        row[f'{key}_GW'] = g.loc[idx, 'p_nom_opt'].sum() / 1e3
        row[f'{key}_added_GW'] = (g.loc[idx, 'p_nom_opt'] - g.loc[idx, 'p_nom']).sum() / 1e3
    re = g.index[g.carrier.isin(['solar', 'wind'])]
    pot = (n.generators_t.p_max_pu.reindex(columns=re).fillna(g.loc[re, 'p_max_pu']) * g.loc[re, 'p_nom_opt'])
    gen = P.reindex(columns=re).fillna(0)
    row['RE_gen_TWh'] = gen.sum().sum() / 1e6
    row['curtail_TWh'] = (pot - gen).clip(lower=0).sum().sum() / 1e6
    row['curtail_%'] = 100 * row['curtail_TWh'] / max(pot.sum().sum() / 1e6, 1e-9)
    st = n.stores[n.stores.index.str.endswith('_ESS')]
    row['ESS_GWh'] = st.e_nom_opt.sum() / 1e3
    row['ESS_added_GWh'] = (st.e_nom_opt - st.e_nom).sum() / 1e3
    ch = n.links.index[n.links.index.str.endswith('_ESS_Ch')]
    row['ESS_GW'] = n.links.loc[ch, 'p_nom_opt'].sum() / 1e3
    row['ESS_charge_TWh'] = n.links_t.p0[ch].clip(lower=0).sum().sum() / 1e6
    v2g = n.links.index[n.links.index.str.endswith('_EV_V2G')]
    row['V2G_TWh'] = n.links_t.p0[v2g].clip(lower=0).sum().sum() / 1e6
    sl = [x for x in g.index if 'slack' in x.lower()]
    row['slack_TWh'] = P.reindex(columns=sl).fillna(0).sum().sum() / 1e6
    hf = [x for x in g.index if 'H_Fallback' in x]
    row['heat_fallback_TWh'] = P.reindex(columns=hf).fillna(0).sum().sum() / 1e6
    lng = g.index[g.carrier == 'gas']
    row['gas_gen_TWh'] = P[lng].sum().sum() / 1e6
    gc = n.global_constraints
    if 'CO2_limit' in gc.index and 'mu' in gc.columns:
        row['CO2_shadow_price'] = float(-gc.at['CO2_limit', 'mu'])
    mp = n.buses_t.marginal_price[[b for b in n.buses.index if b.endswith('_EL')]]
    row['EL_price_mean'] = float(mp.mean().mean())
    row['EL_hours_price0'] = float((mp <= 1).sum().mean())
    return row


def arbitrage_value(price, P, eta_c, eta_d, mc=1.0):
    import linopy
    import xarray as xr
    T = len(price)
    md = linopy.Model()
    t = pd.RangeIndex(T, name='t')
    c = md.add_variables(0, P, coords=[t], name='c')
    d = md.add_variables(0, P, coords=[t], name='d')
    e = md.add_variables(0, 1.0, coords=[t], name='e')
    md.add_constraints(e - e.roll(t=1) - eta_c * c + d == 0, name='bal')
    pr = xr.DataArray(price, coords=[t])
    md.add_objective(((pr + mc) * c + (mc - eta_d * pr) * d).sum())
    md.solve(solver_name='cplex', lpmethod=4, solutiontype=2, threads=8)
    return -float(md.objective.value)


def estimate_breakeven(n, out_dir):
    """기준 결과 전력가격으로 ESS 1MWh(저장시간별)의 연간 차익가치/연간비용 비율 → 진입 배율 k*"""
    rows = []
    for s in n.stores.index[n.stores.index.str.endswith('_ESS')]:
        ch, ds = s + '_Ch', s + '_Disch'
        price = n.buses_t.marginal_price[n.links.at[ch, 'bus0']].values
        ec, ed = float(n.links.at[ch, 'efficiency']), float(n.links.at[ds, 'efficiency'])
        cE = float(n.stores.at[s, 'capital_cost'])
        cP = float(n.links.at[ch, 'capital_cost']) + float(n.links.at[ds, 'capital_cost'])
        for h in DURATIONS_H:
            P = 1.0 / (h * ec)
            v = arbitrage_value(price, P, ec, ed)
            cost = cE + cP * P
            rows.append({'store': s, 'duration_h': h, 'value': v, 'cost': cost, 'ratio': v / cost})
    df = pd.DataFrame(rows)
    df.to_csv(os.path.join(out_dir, 'breakeven_estimate.csv'), index=False, encoding='utf-8-sig')
    best = df.loc[df.groupby('store').ratio.idxmax()]
    print("\n[손익분기 추정] 지역별 최적 저장시간과 가치/비용 비율")
    print(best.to_string(index=False))
    return float(best.ratio.max()), df


def run_case(input_data, mult, label, out_dir):
    print(f"\n{'=' * 80}\n[ESS 비용 민감도] {label}: ESS 비용 배율 = {mult:.3f}\n{'=' * 80}", flush=True)
    data = scale_ess_cost(input_data, mult)
    t0 = time.time()
    n = m.create_network(data)
    ok = bool(m.optimize_network(n))
    elapsed = time.time() - t0
    case_dir = os.path.join(out_dir, label)
    if ok:
        m.save_results(n, subdir=case_dir)
    row = summarize(n, label, mult, ok, elapsed)
    print(f"[ESS 비용 민감도] {label} 완료 ({row['elapsed_min']}분): " + json.dumps(
        {k: (round(v, 2) if isinstance(v, float) else v) for k, v in row.items()}, ensure_ascii=False), flush=True)
    return n, row


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--multipliers', type=float, nargs='*', default=None)
    args = ap.parse_args()

    out_dir = os.path.join('results', f"ESS_cost_sweep_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    os.makedirs(out_dir, exist_ok=True)
    summary_path = os.path.join(out_dir, 'sweep_summary.csv')

    m.ensure_integrated_input()
    input_data = m.read_input_data(m.INPUT_FILE)
    input_data = m.standardize_bus_names_in_input(input_data)
    if 'generators' in input_data and not input_data['generators'].empty:
        input_data['generators']['committable'] = False

    rows = []

    def flush():
        pd.DataFrame(rows).to_csv(summary_path, index=False, encoding='utf-8-sig')

    if args.multipliers:
        mults = [(f"mult_{x:.3f}", x) for x in args.multipliers]
    else:
        n0, r0 = run_case(input_data, 1.0, 'base_1.000', out_dir)
        rows.append(r0); flush()
        if not r0['ok']:
            print("기준 실행 실패 — 중단합니다.")
            return
        k_star, _ = estimate_breakeven(n0, out_dir)
        k_star = min(k_star, 1.0)
        print(f"\n[손익분기 추정] 진입 배율 k* = {k_star:.3f} (현재 ESS 비용의 {100 * k_star:.1f}%)", flush=True)
        with open(os.path.join(out_dir, 'breakeven.json'), 'w', encoding='utf-8') as f:
            json.dump({'k_star': k_star}, f)
        del n0
        mults = [(f"kstar_minus{int(s * 100):02d}_{k_star * (1 - s):.3f}", k_star * (1 - s)) for s in STEPS]

    for label, mult in mults:
        try:
            _, r = run_case(input_data, mult, label, out_dir)
        except Exception as e:
            print(f"[ESS 비용 민감도] {label} 오류: {e}")
            r = {'case': label, 'ess_cost_mult': mult, 'ok': False, 'error': str(e)}
        rows.append(r); flush()

    print("\n[ESS 비용 민감도] 전체 요약")
    print(pd.DataFrame(rows).round(3).to_string(index=False))
    print(f"\n요약 파일: {summary_path}")


if __name__ == '__main__':
    main()
