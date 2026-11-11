"""
송전망 용량과 slack 발생의 관계 연속 분석 (interface 입력조건 그대로, 송전망만 변경)

uniform_{p} : 모든 AC 선로(s_nom)·HVDC(p_nom)를 p% 일괄 증설
uniform_900 : 사실상 송전 제약 없음(10배) 참고 케이스

사용법:
  python run_grid_slack_sweep.py [--pcts 0 10 20 30 900]
"""
import argparse
import os
import time
from datetime import datetime

import pandas as pd

import PyPSA_GUI as m
from run_ess_cost_sweep import summarize
from run_transmission_sweep import apply_uniform, corridor_table, hvdc_links

SLACK_KEYS = ('slack', 'fallback', 'failsafe')


def slack_by_region(n, label):
    g, P = n.generators, n.generators_t.p
    w = n.snapshot_weightings.generators
    rows = []
    for gen in g.index:
        if not any(k in gen.lower() for k in SLACK_KEYS):
            continue
        p = P[gen] if gen in P.columns else pd.Series(0.0, index=n.snapshots)
        if p.max() < 1:
            continue
        rows.append({'case': label, 'generator': gen, 'bus': g.at[gen, 'bus'],
                     'TWh': float((p * w).sum()) / 1e6, 'max_MW': float(p.max()),
                     'hours': float(w[p > 1].sum())})
    return rows


def run_case(input_data, pct, out_dir):
    label = f"uniform_{pct:02d}"
    print(f"\n{'=' * 80}\n[연속 분석] {label}\n{'=' * 80}", flush=True)
    t0 = time.time()
    n = m.create_network(input_data)
    base_lines = n.lines['s_nom'].copy()
    base_hvdc = n.links.loc[hvdc_links(n), 'p_nom'].copy()
    if pct:
        apply_uniform(n, pct)
    ok = bool(m.optimize_network(n))
    row = summarize(n, label, 1.0, ok, time.time() - t0)
    row['grid_pct'] = pct
    corr, sl = [], []
    if ok:
        corr = corridor_table(n, base_lines, base_hvdc, label)
        sl = slack_by_region(n, label)
        row['slack_el_TWh'] = sum(r['TWh'] for r in sl if r['bus'].endswith('_EL'))
        row['grid_added_GW'] = pd.DataFrame(corr).added_MW.sum() / 1e3
        row['grid_annual_cost_bn'] = pd.DataFrame(corr).annual_cost_bn.sum()
        case_dir = os.path.join(out_dir, label)
        os.makedirs(case_dir, exist_ok=True)
        n.export_to_netcdf(os.path.join(case_dir, f'{label}.nc'))
    print(f"[연속 분석] {label} 완료 ({row['elapsed_min']}분) 전력 slack={row.get('slack_el_TWh', 0):.2f}TWh "
          f"열 fallback={row.get('heat_fallback_TWh', 0):.2f}TWh", flush=True)
    return row, corr, sl


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--pcts', nargs='+', type=int, default=[0, 10, 20, 30, 900])
    args = ap.parse_args()

    out_dir = os.path.join('results', f"grid_slack_sweep_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    os.makedirs(out_dir, exist_ok=True)

    m.ensure_integrated_input()
    input_data = m.read_input_data(m.INPUT_FILE)
    input_data = m.standardize_bus_names_in_input(input_data)
    if 'generators' in input_data and not input_data['generators'].empty:
        input_data['generators']['committable'] = False

    rows, corrs, slacks = [], [], []
    for pct in args.pcts:
        try:
            row, corr, sl = run_case(input_data, pct, out_dir)
            rows.append(row); corrs += corr; slacks += sl
        except Exception as e:
            print(f"[연속 분석] uniform_{pct:02d} 오류: {e}", flush=True)
            rows.append({'case': f"uniform_{pct:02d}", 'ok': False, 'error': str(e), 'grid_pct': pct})
        pd.DataFrame(rows).to_csv(os.path.join(out_dir, 'sweep_summary.csv'), index=False, encoding='utf-8-sig')
        pd.DataFrame(corrs).to_csv(os.path.join(out_dir, 'corridor_congestion.csv'), index=False, encoding='utf-8-sig')
        pd.DataFrame(slacks).to_csv(os.path.join(out_dir, 'slack_by_region.csv'), index=False, encoding='utf-8-sig')

    print(f"\n요약 파일: {out_dir}")


if __name__ == '__main__':
    main()
