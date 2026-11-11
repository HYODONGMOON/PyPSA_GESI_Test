"""
수요 감축 민감도 연속 분석 (interface 입력조건 그대로, 전력 수요만 메모리 상에서 축소)

세트 1 (EL)    : {지역}_Demand_EL  -5%, -10%, -15%
세트 2 (DCFab) : {지역}_Demand_DC  -25%, -50%
                 {지역}_Demand_Fab -25%, -50%
                 DC·Fab 동시       -25%, -50%

사용법:
  python run_demand_sweep.py [--set EL DCFab] [--only 케이스 ...]
"""
import argparse
import os
import re
import time
from datetime import datetime

import pandas as pd

import PyPSA_GUI as m
from run_ess_cost_sweep import summarize
from run_regional_cost_sweep import regional_capacity

DEMAND_RE = re.compile(r'^[A-Z]{3}_Demand_(EL|DC|Fab)$')


def target_loads(n, kinds):
    return [l for l in n.loads.index
            if (mt := DEMAND_RE.match(l.strip())) and mt.group(1) in kinds]


def el_demand_twh(n, kinds=('EL', 'DC', 'Fab')):
    w = n.snapshot_weightings.generators
    out = {}
    for k in kinds:
        idx = target_loads(n, (k,))
        p = n.get_switchable_as_dense('Load', 'p_set')[idx]
        out[f'demand_{k}_TWh'] = float(p.mul(w, axis=0).sum().sum()) / 1e6
    return out


def scale_demand(n, kinds, factor):
    idx = target_loads(n, kinds)
    n.loads.loc[idx, 'p_set'] *= factor
    ts = [l for l in idx if l in n.loads_t.p_set.columns]
    n.loads_t.p_set[ts] *= factor
    print(f"[수요 조정] {'/'.join(kinds)} × {factor:.2f} ({len(idx)}개 부하, 시계열 {len(ts)}개)", flush=True)


def mix_rows(n, label):
    g, P = n.generators, n.generators_t.p
    w = n.snapshot_weightings.generators
    rows = []
    gen = P.mul(w, axis=0).sum()
    for car, idx in g.groupby('carrier').groups.items():
        rows.append({'case': label, 'component': 'generator', 'carrier': car,
                     'capacity_GW': g.loc[idx, 'p_nom_opt'].sum() / 1e3,
                     'added_GW': (g.loc[idx, 'p_nom_opt'] - g.loc[idx, 'p_nom']).sum() / 1e3,
                     'energy_TWh': gen.reindex(idx).fillna(0).sum() / 1e6})
    L = n.links
    out = n.links_t.p1.mul(w, axis=0).sum().mul(-1)
    for car, idx in L[L.bus1.str.endswith('_EL')].groupby('carrier').groups.items():
        rows.append({'case': label, 'component': 'link_to_EL', 'carrier': car,
                     'capacity_GW': L.loc[idx, 'p_nom_opt'].sum() / 1e3,
                     'added_GW': (L.loc[idx, 'p_nom_opt'] - L.loc[idx, 'p_nom']).sum() / 1e3,
                     'energy_TWh': out.reindex(idx).fillna(0).sum() / 1e6})
    return rows


def run_case(input_data, label, out_dir, kinds, factor):
    print(f"\n{'=' * 80}\n[연속 분석] {label}\n{'=' * 80}", flush=True)
    t0 = time.time()
    n = m.create_network(input_data)
    if kinds:
        scale_demand(n, kinds, factor)
    dem = el_demand_twh(n)
    ok = bool(m.optimize_network(n))
    row = summarize(n, label, 1.0, ok, time.time() - t0)
    row.update({'kinds': '+'.join(kinds) if kinds else 'base', 'factor': factor, **dem})
    mix, reg = [], []
    if ok:
        mix = mix_rows(n, label)
        reg = regional_capacity(n, label)
        case_dir = os.path.join(out_dir, label)
        os.makedirs(case_dir, exist_ok=True)
        n.export_to_netcdf(os.path.join(case_dir, f'{label}.nc'))
    print(f"[연속 분석] {label} 완료 ({row['elapsed_min']}분) PV={row.get('PV_GW', 0):.1f}GW "
          f"WT={row.get('WT_GW', 0):.1f}GW slack={row.get('slack_TWh', 0):.2f}TWh "
          f"출력제한={row.get('curtail_%', 0):.2f}%", flush=True)
    return row, mix, reg


def build_cases(sets):
    cases = [('base', (), 1.0)]
    if 'EL' in sets:
        cases += [(f'EL_m{p:02d}', ('EL',), 1 - p / 100) for p in (5, 10, 15)]
    if 'DCFab' in sets:
        for kinds, tag in ((('DC',), 'DC'), (('Fab',), 'Fab'), (('DC', 'Fab'), 'DCFab')):
            cases += [(f'{tag}_m{p:02d}', kinds, 1 - p / 100) for p in (25, 50)]
    return cases


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--set', nargs='+', default=['EL', 'DCFab'], choices=['EL', 'DCFab'])
    ap.add_argument('--only', nargs='+')
    args = ap.parse_args()

    cases = build_cases(args.set)
    if args.only:
        cases = [c for c in cases if c[0] in args.only]

    out_dir = os.path.join('results', f"demand_sweep_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    os.makedirs(out_dir, exist_ok=True)

    m.ensure_integrated_input()
    input_data = m.read_input_data(m.INPUT_FILE)
    input_data = m.standardize_bus_names_in_input(input_data)
    if 'generators' in input_data and not input_data['generators'].empty:
        input_data['generators']['committable'] = False

    rows, mixes, regs = [], [], []
    for label, kinds, factor in cases:
        try:
            row, mix, reg = run_case(input_data, label, out_dir, kinds, factor)
            rows.append(row); mixes += mix; regs += reg
        except Exception as e:
            print(f"[연속 분석] {label} 오류: {e}", flush=True)
            rows.append({'case': label, 'ok': False, 'error': str(e)})
        pd.DataFrame(rows).to_csv(os.path.join(out_dir, 'sweep_summary.csv'), index=False, encoding='utf-8-sig')
        pd.DataFrame(mixes).to_csv(os.path.join(out_dir, 'mix_by_carrier.csv'), index=False, encoding='utf-8-sig')
        pd.DataFrame(regs).to_csv(os.path.join(out_dir, 'regional_capacity.csv'), index=False, encoding='utf-8-sig')

    print(f"\n요약 파일: {out_dir}")


if __name__ == '__main__':
    main()
