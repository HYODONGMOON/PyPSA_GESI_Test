"""
출력제한 페널티 / 풍력 하한 + ESS 비용 연속 분석

interface/통합 입력 파일은 수정하지 않고 메모리 상에서만 조건을 바꿔 연속 실행한다.

A. 출력제한 페널티 c (천원/MWh):
   PV·WT에 capital_cost += c × Σ_t p_max_pu, marginal_cost -= c 를 적용
   → 목적함수에 c × (가용량 − 발전량) = c × 출력제한량 이 더해지는 것과 동일
B. 풍력 하한 유지 + ESS 비용 배율:
   지역별 WT p_nom_min = 기준 결과의 WT 용량, ESS(Store + Ch/Disch Link) capital_cost × 배율

C. 출력제한 페널티 30 + ESS 비용 배율
D. 원자력 최소출력(p_min_pu) 완화
E. 원자력 최소출력 0.6 + 출력제한 페널티 30

사용법:
  python run_flex_sweep.py --base-nc <기준 결과 .nc> [--set AB|CDE]
"""
import argparse
import glob
import os
import time
from datetime import datetime

import pandas as pd
import pypsa

import PyPSA_GUI as m
from run_ess_cost_sweep import scale_ess_cost, summarize

PENALTIES = (10.0, 30.0, 50.0, 100.0)
ESS_MULTS = (1.0, 0.382, 0.305, 0.267)


def apply_curtailment_penalty(n, c):
    re = n.generators.index[n.generators.carrier.isin(['solar', 'wind'])]
    w = n.snapshot_weightings.generators
    pmax = n.generators_t.p_max_pu.reindex(columns=re).fillna(n.generators.loc[re, 'p_max_pu'])
    hours = pmax.mul(w, axis=0).sum()
    ext = re[n.generators.loc[re, 'p_nom_extendable'].astype(bool)]
    n.generators.loc[ext, 'capital_cost'] += c * hours[ext]
    n.generators.loc[re, 'marginal_cost'] -= c
    print(f"[출력제한 페널티] c={c} 천원/MWh 적용: 발전기 {len(re)}개 "
          f"(PV 평균 가산 {c * hours[re[n.generators.loc[re, 'carrier'] == 'solar']].mean():,.0f}, "
          f"WT 평균 가산 {c * hours[re[n.generators.loc[re, 'carrier'] == 'wind']].mean():,.0f} 천원/MW·년)")
    return re, ext, hours


def penalty_in_objective(n, c, re, ext, hours):
    """목적함수에 포함된 페널티 항. PyPSA는 확장 설비의 기존 용량 고정비를 상수로 빼므로 기존 용량분은 제외된다."""
    g = n.generators
    w = n.snapshot_weightings.generators
    cap_term = (c * hours[ext] * (g.loc[ext, 'p_nom_opt'] - g.loc[ext, 'p_nom'])).sum()
    dispatch = n.generators_t.p[re].mul(w, axis=0).sum().sum()
    return cap_term - c * dispatch


def apply_wind_floor(n, floor):
    for gname, cap in floor.items():
        if gname in n.generators.index:
            cur = float(n.generators.at[gname, 'p_nom_min'] or 0)
            n.generators.at[gname, 'p_nom_min'] = max(cur, cap)
            mx = float(n.generators.at[gname, 'p_nom_max'])
            if mx < cap:
                n.generators.at[gname, 'p_nom_max'] = cap
    print(f"[풍력 하한] {len(floor)}개 발전기, 합계 {sum(floor.values()) / 1e3:.1f} GW")


def apply_nuclear_pmin(n, pmin):
    nu = n.generators.index[n.generators.carrier == 'nuclear']
    n.generators.loc[nu, 'p_min_pu'] = n.generators.loc[nu, 'p_min_pu'].clip(upper=pmin)
    ts = n.generators_t.p_min_pu
    cols = [c for c in nu if c in ts.columns]
    if cols:
        ts[cols] = ts[cols].clip(upper=pmin)
    print(f"[원자력 최소출력] {len(nu)}기, p_min_pu ≤ {pmin}")


def run_case(input_data, label, out_dir, ess_mult=1.0, penalty=0.0, wind_floor=None, nuclear_pmin=None):
    print(f"\n{'=' * 80}\n[연속 분석] {label}: ESS배율={ess_mult}, 출력제한페널티={penalty}, "
          f"풍력하한={'적용' if wind_floor else '없음'}, 원자력최소출력={nuclear_pmin or '기존'}\n{'=' * 80}",
          flush=True)
    data = scale_ess_cost(input_data, ess_mult) if ess_mult != 1.0 else input_data
    t0 = time.time()
    n = m.create_network(data)
    pen_info = apply_curtailment_penalty(n, penalty) if penalty else None
    if wind_floor:
        apply_wind_floor(n, wind_floor)
    if nuclear_pmin is not None:
        apply_nuclear_pmin(n, nuclear_pmin)
    ok = bool(m.optimize_network(n))
    elapsed = time.time() - t0
    row = summarize(n, label, ess_mult, ok, elapsed)
    row.update({'curtail_penalty': penalty, 'wind_floor': bool(wind_floor), 'nuclear_pmin': nuclear_pmin})
    if ok:
        row['objective_excl_penalty'] = row['objective'] - (
            penalty_in_objective(n, penalty, *pen_info) if pen_info else 0.0)
        case_dir = os.path.join(out_dir, label)
        os.makedirs(case_dir, exist_ok=True)
        n.export_to_netcdf(os.path.join(case_dir, f'{label}.nc'))
    print(f"[연속 분석] {label} 완료 ({row['elapsed_min']}분) PV={row.get('PV_GW', 0):.1f}GW "
          f"WT={row.get('WT_GW', 0):.1f}GW 출력제한={row.get('curtail_%', 0):.2f}% "
          f"ESS={row.get('ESS_GWh', 0):.0f}GWh", flush=True)
    return row


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--base-nc', required=True)
    ap.add_argument('--set', choices=['AB', 'CDE'], default='AB')
    args = ap.parse_args()

    base = pypsa.Network(args.base_nc)
    wg = base.generators.index[base.generators.carrier == 'wind']
    wind_floor = base.generators.loc[wg, 'p_nom_opt'].to_dict()
    del base

    out_dir = os.path.join('results', f"flex_sweep_{args.set}_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    os.makedirs(out_dir, exist_ok=True)
    summary_path = os.path.join(out_dir, 'sweep_summary.csv')

    m.ensure_integrated_input()
    input_data = m.read_input_data(m.INPUT_FILE)
    input_data = m.standardize_bus_names_in_input(input_data)
    if 'generators' in input_data and not input_data['generators'].empty:
        input_data['generators']['committable'] = False

    if args.set == 'AB':
        cases = [(f"A_penalty_{int(c):03d}", dict(penalty=c)) for c in PENALTIES]
        cases += [(f"B_windfloor_ess{x:.3f}", dict(ess_mult=x, wind_floor=wind_floor)) for x in ESS_MULTS]
    else:
        cases = [(f"C_penalty030_ess{x:.3f}", dict(penalty=30.0, ess_mult=x)) for x in ESS_MULTS[1:]]
        cases += [(f"D_nucpmin_{p:.1f}", dict(nuclear_pmin=p)) for p in (0.6, 0.4)]
        cases += [("E_nucpmin_0.6_penalty030", dict(nuclear_pmin=0.6, penalty=30.0))]

    rows = []
    for label, kw in cases:
        try:
            rows.append(run_case(input_data, label, out_dir, **kw))
        except Exception as e:
            print(f"[연속 분석] {label} 오류: {e}")
            rows.append({'case': label, 'ok': False, 'error': str(e), **{k: v for k, v in kw.items() if k != 'wind_floor'}})
        pd.DataFrame(rows).to_csv(summary_path, index=False, encoding='utf-8-sig')

    print("\n[연속 분석] 전체 요약")
    print(pd.DataFrame(rows).round(3).to_string(index=False))
    print(f"\n요약 파일: {summary_path}")


if __name__ == '__main__':
    main()
