"""
반도체 수요 재배치(WP2) 추가 민감도 연속 분석 (저장된 네트워크에서 조건만 바꿔 재최적화)

1. REFWP2      : 재배치 + REF 재생 상한(p_nom_max를 REF 결과에서 가져옴)
2. {OPEN,WP2}cap     : 지역 상한 = max(기술적 잠재량 × 20%, 기존 설비)
3. {OPEN,WP2}capgrid : 2 + 필요 구간 송전 확장(건설비 × 1)

입력 : wp2_compare_nc/{OPEN_2035, WP2fix_2035, REF_2035, OPEN_2040, WP2_2040, REF_2040}.nc
출력 : wp2_compare_nc/{케이스}_{연도}.nc, wp2_sensitivity.log

사용법:
  python run_wp2_sensitivity.py [--only 케이스_연도 ...]
"""
import argparse
import logging
import os
import time

import pypsa

import PyPSA_GUI as m
from run_regional_cost_sweep import apply_potential_cap
from run_transmission_sweep import apply_need_based

logging.disable(logging.WARNING)
SRC = 'wp2_compare_nc'
BASE = {('OPEN', '2035'): 'OPEN_2035', ('WP2', '2035'): 'WP2fix_2035',
        ('OPEN', '2040'): 'OPEN_2040', ('WP2', '2040'): 'WP2_2040'}


def build(case, year):
    if case == 'REFWP2':
        n = pypsa.Network(f'{SRC}/{BASE["WP2", year]}.nc')
        ref = pypsa.Network(f'{SRC}/REF_{year}.nc')
        n.generators['p_nom_max'] = ref.generators.p_nom_max.reindex(n.generators.index).fillna(
            n.generators.p_nom_max)
        print(f'[조건] REF 상한 적용: PV+풍력 상한 합계 '
              f'{n.generators[n.generators.carrier.isin(["solar", "wind"])].p_nom_max.sum() / 1e3:.1f} GW', flush=True)
        return n
    alloc = 'OPEN' if case.startswith('OPEN') else 'WP2'
    n = pypsa.Network(f'{SRC}/{BASE[alloc, year]}.nc')
    apply_potential_cap(n, 0.2)
    if case.endswith('grid'):
        apply_need_based(n, 1.0)
    return n


def solved(n):
    ch = [l for l in n.links.index if l.endswith('_ESS_Ch')]
    return getattr(n, 'objective', None) is not None and all(c in n.links_t.p0.columns for c in ch)


def run(case, year):
    label = f'{case}_{year}'
    print(f"\n{'=' * 80}\n[연속 분석] {label}\n{'=' * 80}", flush=True)
    t0 = time.time()
    for tol in (None, '1e-5'):
        if tol:
            os.environ['PYPSA_CPLEX_CONVERGETOL'] = tol
            print(f'[재시도] convergetol={tol}', flush=True)
        n = build(case, year)
        n.objective = None
        n.links_t.p0 = n.links_t.p0.iloc[:, :0]
        m.optimize_network(n)
        if solved(n):
            break
    os.environ.pop('PYPSA_CPLEX_CONVERGETOL', None)
    if not solved(n):
        print(f'[연속 분석] {label} 실패', flush=True)
        return
    n.export_to_netcdf(f'{SRC}/{label}.nc')
    print(f'[연속 분석] {label} 완료 ({(time.time() - t0) / 60:.1f}분) objective={n.objective / 1e9:.2f}조', flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--only', nargs='+')
    args = ap.parse_args()
    cases = [(c, y) for y in ('2040', '2035')
             for c in ('REFWP2', 'OPENcap', 'WP2cap', 'OPENcapgrid', 'WP2capgrid')]
    if args.only:
        cases = [c for c in cases if f'{c[0]}_{c[1]}' in args.only]
    for c, y in cases:
        try:
            run(c, y)
        except Exception as e:
            print(f'[연속 분석] {c}_{y} 오류: {e}', flush=True)


if __name__ == '__main__':
    main()
