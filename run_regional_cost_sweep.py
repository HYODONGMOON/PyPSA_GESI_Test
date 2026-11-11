"""
지역별 재생에너지 비용 차등 연속 분석

interface/통합 입력 파일은 수정하지 않고 메모리 상에서만 PV·WT capital_cost를 지역별로 조정한다.

지역 배율 m_r = k × (P_r / mean(P))^(−α)
  P_r : 지역별 기술적 잠재량 (interface 'RE잠재량_REF' 시트)
        PV = KIER 신재생에너지자원지도 PV 기술적 잠재량
        WT = KIER 육상풍력 기술적 잠재량 + 해양공간계획 유보해역 기반 해상풍력 잠재량
  k   : 기존 설비용량 가중평균 배율 = 1 (현재 전국 평균 비용 유지)

--set potential : 절대 잠재량 기준 배율, α = --alphas
--set mix       : 상한(잠재량 × 비율) / 1인당 잠재량 기준 배율 / 수도권 가중치 케이스

사용법:
  python run_regional_cost_sweep.py [--set potential|mix] [--alphas 0.25 0.5 0.75]
"""
import argparse
import os
import time
from datetime import datetime

import pandas as pd

import PyPSA_GUI as m
from run_ess_cost_sweep import summarize

TECH_POTENTIAL_GW = pd.DataFrame({
    'PV': {'BSN': 38, 'CBD': 118, 'CND': 237, 'GBD': 283, 'GGD': 254, 'GND': 172,
           'GWD': 131, 'ICN': 37, 'JBD': 185, 'JJD': 60, 'JND': 292},
    'WT': {'BSN': 6 + 3.48, 'CBD': 25 + 0.0, 'CND': 37 + 5.95, 'GBD': 72 + 11.75, 'GGD': 44 + 0.40,
           'GND': 33 + 4.93, 'GWD': 48 + 8.48, 'ICN': 4 + 13.33, 'JBD': 29 + 3.74, 'JJD': 9 + 15.10,
           'JND': 45 + 56.33},
})

# 2024년 주민등록인구 근사치 (백만 명)
POPULATION_M = pd.Series({'BSN': 3.28 + 1.10, 'CBD': 1.59, 'CND': 1.44 + 0.39 + 2.13, 'GBD': 2.37 + 2.55,
                          'GGD': 9.35 + 13.69, 'GND': 3.25, 'GWD': 1.52, 'ICN': 3.00, 'JBD': 1.75,
                          'JJD': 0.67, 'JND': 1.41 + 1.79})
METRO = ('GGD', 'ICN')


def regional_multipliers(n, alpha, index='potential'):
    g = n.generators
    mult = {}
    for tech in ('PV', 'WT'):
        pot = TECH_POTENTIAL_GW[tech]
        if index == 'per_capita':
            pot = pot / POPULATION_M
        names = [f'{r}_{tech}' for r in pot.index]
        existing = g.loc[names, 'p_nom'].set_axis(pot.index)
        raw = (pot / pot.mean()) ** (-alpha)
        mult[tech] = raw / ((existing * raw).sum() / existing.sum())
    return pd.DataFrame(mult)


def metro_multipliers(weight):
    idx = TECH_POTENTIAL_GW.index
    x = pd.Series([weight if r in METRO else 1.0 for r in idx], index=idx)
    return pd.DataFrame({'PV': x, 'WT': x})


def apply_cost_multipliers(n, mult, tag):
    for tech in mult.columns:
        for r, x in mult[tech].items():
            for name in (f'{r}_{tech}', f'{r}_{tech}_RE'):
                if name in n.generators.index:
                    n.generators.at[name, 'capital_cost'] *= x
    print(f"[지역 비용 차등] {tag}\n{mult.round(3).to_string()}", flush=True)
    return mult


def apply_regional_cost(n, alpha, index='potential'):
    return apply_cost_multipliers(n, regional_multipliers(n, alpha, index), f"{index} α={alpha}")


def apply_potential_cap(n, share):
    g = n.generators
    for tech in ('PV', 'WT'):
        for r, pot in TECH_POTENTIAL_GW[tech].items():
            name = f'{r}_{tech}'
            cap = max(share * pot * 1e3, float(g.at[name, 'p_nom']))
            cur = pd.to_numeric(g.at[name, 'p_nom_max'], errors='coerce')
            g.at[name, 'p_nom_max'] = cap if pd.isna(cur) else min(float(cur), cap)
    caps = g.loc[[f'{r}_{t}' for t in ('PV', 'WT') for r in TECH_POTENTIAL_GW.index], 'p_nom_max'] / 1e3
    print(f"[잠재량 상한] 기술적 잠재량 × {share:.0%}: PV 합계 {caps.filter(like='_PV').sum():.0f} GW, "
          f"WT 합계 {caps.filter(like='_WT').sum():.0f} GW", flush=True)


def regional_capacity(n, label):
    g = n.generators
    rows = []
    for tech in ('PV', 'WT'):
        for r, pot in TECH_POTENTIAL_GW[tech].items():
            name = f'{r}_{tech}'
            rows.append({'case': label, 'region': r, 'tech': tech, 'existing_GW': g.at[name, 'p_nom'] / 1e3,
                         'opt_GW': g.at[name, 'p_nom_opt'] / 1e3, 'potential_GW': pot,
                         'use_%': g.at[name, 'p_nom_opt'] / 1e3 / pot * 100})
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--set', choices=['potential', 'mix'], default='potential')
    ap.add_argument('--alphas', type=float, nargs='+', default=[0.25, 0.5, 0.75])
    ap.add_argument('--only', nargs='+', help='지정한 케이스 이름만 실행')
    args = ap.parse_args()

    if args.set == 'potential':
        cases = [(f"regcost_alpha{a:.2f}", [('cost', a, 'potential')]) for a in args.alphas]
    else:
        cases = [
            ("cap20_uniform", [('cap', 0.2)]),
            ("cap20_regcost_alpha0.25", [('cap', 0.2), ('cost', 0.25, 'potential')]),
            ("percap_alpha0.25", [('cost', 0.25, 'per_capita')]),
            ("percap_alpha0.50", [('cost', 0.5, 'per_capita')]),
            ("metro_w1.3", [('metro', 1.3)]),
            ("metro_w1.5", [('metro', 1.5)]),
            ("cap20_percap_alpha0.25", [('cap', 0.2), ('cost', 0.25, 'per_capita')]),
        ]
    if args.only:
        cases = [c for c in cases if c[0] in args.only]
    if args.only:
        cases = [c for c in cases if c[0] in args.only]

    out_dir = os.path.join('results', f"regional_cost_sweep_{args.set}_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    os.makedirs(out_dir, exist_ok=True)

    m.ensure_integrated_input()
    input_data = m.read_input_data(m.INPUT_FILE)
    input_data = m.standardize_bus_names_in_input(input_data)
    if 'generators' in input_data and not input_data['generators'].empty:
        input_data['generators']['committable'] = False

    rows, reg_rows, mults = [], [], []
    for label, mods in cases:
        print(f"\n{'=' * 80}\n[연속 분석] {label}\n{'=' * 80}", flush=True)
        try:
            t0 = time.time()
            n = m.create_network(input_data)
            for mod in mods:
                if mod[0] == 'cap':
                    apply_potential_cap(n, mod[1])
                elif mod[0] == 'cost':
                    mults.append(apply_regional_cost(n, mod[1], mod[2]).assign(case=label))
                elif mod[0] == 'metro':
                    mults.append(apply_cost_multipliers(n, metro_multipliers(mod[1]), f"수도권 ×{mod[1]}")
                                 .assign(case=label))
            ok = bool(m.optimize_network(n))
            row = summarize(n, label, 1.0, ok, time.time() - t0)
            row['modifiers'] = str(mods)
            rows.append(row)
            if ok:
                reg_rows += regional_capacity(n, label)
                case_dir = os.path.join(out_dir, label)
                os.makedirs(case_dir, exist_ok=True)
                n.export_to_netcdf(os.path.join(case_dir, f'{label}.nc'))
            print(f"[연속 분석] {label} 완료 ({row['elapsed_min']}분) PV={row.get('PV_GW', 0):.1f}GW "
                  f"WT={row.get('WT_GW', 0):.1f}GW 출력제한={row.get('curtail_%', 0):.2f}%", flush=True)
        except Exception as e:
            print(f"[연속 분석] {label} 오류: {e}")
            rows.append({'case': label, 'ok': False, 'error': str(e), 'modifiers': str(mods)})
        pd.DataFrame(rows).to_csv(os.path.join(out_dir, 'sweep_summary.csv'), index=False, encoding='utf-8-sig')
        pd.DataFrame(reg_rows).to_csv(os.path.join(out_dir, 'regional_capacity.csv'), index=False, encoding='utf-8-sig')
        if mults:
            pd.concat(mults).to_csv(os.path.join(out_dir, 'cost_multipliers.csv'), encoding='utf-8-sig')

    print(f"\n요약 파일: {out_dir}")


if __name__ == '__main__':
    main()
