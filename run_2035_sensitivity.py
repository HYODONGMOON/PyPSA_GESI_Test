"""
2035년 민감도 분석 (interface 현재 입력 = 2035년, 재생 보급 상한 해제 기준)

케이스:
  base        : 기준
  cap20       : 지역별 PV·풍력 상한 = max(기술적 잠재량 × 20%, 기존 설비)
  EL_m05/m10  : {지역}_Demand_EL -5%, -10%
  DC_m25      : {지역}_Demand_DC -25%
  Fab_m25     : {지역}_Demand_Fab -25%
  DCFab_m25/50: DC·Fab 동시 -25%, -50%
  A           : 대표경로 A (잠재량 20% 상한 + DC -50% + Fab -8.4%)
  (MGA 1%, 2%는 run_mga_sweep.py --eps 1 2 로 별도 실행)

사용법:
  python run_2035_sensitivity.py [--only 케이스 ...]
"""
import argparse
import os
import time
from datetime import datetime

import pandas as pd

import PyPSA_GUI as m
from run_demand_sweep import el_demand_twh, mix_rows, scale_demand
from run_ess_cost_sweep import summarize
from run_mga_sweep import solution_assigned
from run_regional_cost_sweep import apply_potential_cap, regional_capacity

TOL_ENV = 'PYPSA_CPLEX_CONVERGETOL'
CASES = [
    ('base', {}),
    ('cap20', {'cap': 0.2}),
    ('EL_m05', {'demand': (('EL',), 0.95)}),
    ('EL_m10', {'demand': (('EL',), 0.90)}),
    ('DC_m25', {'demand': (('DC',), 0.75)}),
    ('Fab_m25', {'demand': (('Fab',), 0.75)}),
    ('DCFab_m25', {'demand': (('DC', 'Fab'), 0.75)}),
    ('DCFab_m50', {'demand': (('DC', 'Fab'), 0.50)}),
    ('A', {'cap': 0.2, 'demand': [(('DC',), 0.50), (('Fab',), 0.916)]}),
]


def build(input_data, mods):
    n = m.create_network(input_data)
    if 'cap' in mods:
        apply_potential_cap(n, mods['cap'])
    dem = mods.get('demand', [])
    for kinds, factor in ([dem] if isinstance(dem, tuple) else dem):
        scale_demand(n, kinds, factor)
    return n


def run_case(input_data, label, mods, out_dir):
    print(f"\n{'=' * 80}\n[2035 민감도] {label}\n{'=' * 80}", flush=True)
    t0 = time.time()
    ok, tol_used = False, None
    for tol in (None, '1e-5'):
        if tol:
            os.environ[TOL_ENV] = tol
            print(f"[2035 민감도] {label}: 수렴 실패 → convergetol {tol} 재시도", flush=True)
        n = build(input_data, mods)
        dem = el_demand_twh(n)
        try:
            ok = bool(m.optimize_network(n)) and solution_assigned(n)
        finally:
            os.environ.pop(TOL_ENV, None)
        if ok:
            tol_used = tol or '1e-6'
            break
    row = summarize(n, label, 1.0, ok, time.time() - t0)
    scales = mods.get('demand', [])
    scales = [scales] if isinstance(scales, tuple) else scales
    row.update({'kinds': ' '.join(f"{'+'.join(k)}×{f:g}" for k, f in scales) or 'base', 'cap_share': mods.get('cap'),
                'convergetol': tol_used, **dem})
    mix, reg = [], []
    if ok:
        mix = mix_rows(n, label)
        reg = regional_capacity(n, label)
        n.export_to_netcdf(os.path.join(out_dir, f'{label}.nc'))
    print(f"[2035 민감도] {label} 완료 ({row['elapsed_min']}분) PV={row.get('PV_GW', 0):.1f}GW "
          f"WT={row.get('WT_GW', 0):.1f}GW slack={row.get('slack_TWh', 0):.2f}TWh "
          f"출력제한={row.get('curtail_%', 0):.2f}%", flush=True)
    return row, mix, reg


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--only', nargs='+')
    args = ap.parse_args()
    cases = [c for c in CASES if not args.only or c[0] in args.only]

    out_dir = os.path.join('results', f"sens2035_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    os.makedirs(out_dir, exist_ok=True)

    m.ensure_integrated_input()
    input_data = m.read_input_data(m.INPUT_FILE)
    input_data = m.standardize_bus_names_in_input(input_data)
    if 'generators' in input_data and not input_data['generators'].empty:
        input_data['generators']['committable'] = False

    rows, mixes, regs = [], [], []
    for label, mods in cases:
        try:
            row, mix, reg = run_case(input_data, label, mods, out_dir)
            rows.append(row); mixes += mix; regs += reg
        except Exception as e:
            print(f"[2035 민감도] {label} 오류: {e}", flush=True)
            rows.append({'case': label, 'ok': False, 'error': str(e)})
        pd.DataFrame(rows).to_csv(os.path.join(out_dir, 'sweep_summary.csv'), index=False, encoding='utf-8-sig')
        pd.DataFrame(mixes).to_csv(os.path.join(out_dir, 'mix_by_carrier.csv'), index=False, encoding='utf-8-sig')
        pd.DataFrame(regs).to_csv(os.path.join(out_dir, 'regional_capacity.csv'), index=False, encoding='utf-8-sig')

    print(f"\n요약 파일: {out_dir}")


if __name__ == '__main__':
    main()
