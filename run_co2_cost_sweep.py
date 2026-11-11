"""
비용 상한 기반 최소 배출 대안 분석 (CO2 제약 사실상 해제, 현재 interface 조건)

interface/통합 입력 파일은 수정하지 않는다. PyPSA_GUI.optimize_network를 그대로 쓰고,
extra_functionality 뒤에 비용 상한·목적함수 교체만 덧붙인다 (run_mga_sweep과 같은 방식).

  - 비용 상한 = PyPSA 목적함수 전체 (기존 분석과 같은 비용 정의, slack 포함)
  - 보고용 시스템비용 C_sys = 목적함수 − slack 비용 전부 (후처리에서만 제외)
    (열/수소 Fallback·전력 slack·EV failsafe 운전 페널티 + slack 설비 투자비)
  - 배출량 E = Σ 발전기 p / 효율 × 배출계수 (CO2_limit GlobalConstraint와 같은 산식,
    CHP 연료는 LNG_Fuel_Supply 발전기로 포함)
  - 대안 단계에서 Fallback·slack 연간 공급량은 발전기별 B0 이하로 제한

단계
  B0      : 비용 최소화 → Z_B0(목적함수), E_B0
  ε 1단계 : Z ≤ (1+ε)·Z_B0 (ε=0이면 상대 허용오차 1e-5) 하에서 E 최소화 → E*
  ε 2단계 : E ≤ E*(1+1e-4) 하에서 원래 목적함수(비용) 최소화

사용법:
  python run_co2_cost_sweep.py [--eps 0 1 3 5 10] [--tag 2040]
"""
import argparse
import json
import os
import time
from datetime import datetime

import pandas as pd
import xarray as xr

import PyPSA_GUI as m
from run_mga_sweep import (COST_SCALE, co2_emissions, cost_expr, fallback_cap, fallback_gens, load_input,
                           penalty_value, power_slack_gens, record, solve, system_cost)
from run_regional_cost_sweep import regional_capacity

E_TOL = 1e-4
E_SCALE = 1e3
BUDGET_SCALE = 1e9


def emission_expr(n, sns):
    """co2_emissions()와 같은 산식의 선형식 (tCO2)"""
    mdl = n.model
    em = n.carriers['co2_emissions']
    em = em[em != 0]
    g = n.generators[n.generators.carrier.isin(em.index)]
    w = n.snapshot_weightings.generators.loc[sns]
    eff = n.get_switchable_as_dense('Generator', 'efficiency', snapshots=sns)[g.index]
    coef = (1.0 / eff).mul(g.carrier.map(em), axis=1).mul(w, axis=0)
    c = xr.DataArray(coef.values, coords={'snapshot': sns, 'Generator': g.index}, dims=['snapshot', 'Generator'])
    p = mdl.variables['Generator-p'].sel(Generator=g.index, snapshot=sns)
    return (p * c).sum()


def aux_gens(n):
    return fallback_gens(n) + power_slack_gens(n)


def aux_ext(n):
    g = n.generators.loc[aux_gens(n)]
    return list(g.index[g.p_nom_extendable.astype(bool) & (g.capital_cost != 0)])


def aux_capex(n):
    g = n.generators.loc[aux_ext(n)]
    return float((g.capital_cost * (g.p_nom_opt - g.p_nom)).sum())


def sys_cost(n):
    """목적함수 − slack·Fallback 운전 페널티 − slack 설비 투자비"""
    return float(n.objective) - penalty_value(n) - aux_capex(n)


def budget_expr(n, sns):
    """PyPSA 목적함수 식 그대로 (slack 페널티·투자비 포함, 기존 설비 상수항 포함)"""
    obj = n.model.objective
    return obj.expression if hasattr(obj, 'expression') else obj


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--eps', type=float, nargs='+', default=[0, 1, 3, 5, 10])
    ap.add_argument('--zero-tol', type=float, default=1e-5)
    ap.add_argument('--tag', default='')
    ap.add_argument('--ref-objective', type=float, help='기준 결과(GUI) 목적함수, B0 일치 확인용 (원)')
    ap.add_argument('--reuse-b0', help='B0.nc가 있는 이전 실행 폴더 (B0 재계산 생략)')
    args = ap.parse_args()

    out_dir = os.path.join('results', f"co2_cost_sweep_{args.tag + '_' if args.tag else ''}"
                                      f"{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    os.makedirs(out_dir, exist_ok=True)
    rows, regs = [], []

    def flush():
        pd.DataFrame(rows).to_csv(os.path.join(out_dir, 'summary.csv'), index=False, encoding='utf-8-sig')
        pd.DataFrame(regs).to_csv(os.path.join(out_dir, 'regional_capacity.csv'), index=False, encoding='utf-8-sig')

    input_data = load_input()

    # ── B0 ──────────────────────────────────────────────────────────────
    print(f"\n{'=' * 80}\n[CO2-비용] B0: 비용 최소화 (현재 interface 조건)\n{'=' * 80}", flush=True)
    t0 = time.time()
    if args.reuse_b0:
        import pypsa
        n = pypsa.Network(os.path.join(args.reuse_b0, 'B0.nc'))
        print(f"[CO2-비용] B0 재사용: {args.reuse_b0}", flush=True)
    else:
        n = m.create_network(input_data)
        if not solve(n):
            print("[CO2-비용] B0 실패 — 중단")
            return
    obj_B0 = float(n.objective)
    pen_B0 = penalty_value(n)
    slack_capex_B0 = aux_capex(n)
    C_B0 = sys_cost(n)
    E_B0 = co2_emissions(n) * 1e6
    w = n.snapshot_weightings.generators
    aux_energy = {g: float(n.generators_t.p[g].mul(w).sum()) for g in aux_gens(n)}
    stat_cost = system_cost(n)[0] - slack_capex_B0
    base = dict(case='B0', eps_pct=0.0, cost_cap=None, C_sys=C_B0, objective_full=obj_B0, penalty_cost=pen_B0,
                slack_capex=slack_capex_B0, objective_increase_pct=0.0, cost_increase_pct=0.0, statistics_cost_check=stat_cost,
                ref_objective=args.ref_objective)
    rows.append(record(n, 'B0', True, t0, base))
    regs += regional_capacity(n, 'B0')
    n.export_to_netcdf(os.path.join(out_dir, 'B0.nc'))
    with open(os.path.join(out_dir, 'B0_reference.json'), 'w', encoding='utf-8') as f:
        json.dump({'objective': obj_B0, 'penalty': pen_B0, 'slack_capex': slack_capex_B0, 'C_B0': C_B0, 'E_B0_t': E_B0,
                   'aux_energy_MWh': aux_energy, 'ref_objective': args.ref_objective}, f, ensure_ascii=False, indent=1)
    msg = f" (기준 결과 {args.ref_objective / 1e9:.5f}조)" if args.ref_objective else ''
    print(f"[CO2-비용] B0 목적함수={obj_B0 / 1e9:.5f}조{msg}, 페널티={pen_B0 / 1e9:.4f}조, "
          f"slack 투자비={slack_capex_B0 / 1e9:.4f}조, "
          f"C_B0={C_B0 / 1e9:.4f}조, CO2={E_B0 / 1e6:.3f}Mt", flush=True)
    flush()
    del n

    # ── ε 대안 ─────────────────────────────────────────────────────────
    for eps in args.eps:
        label = f"CO2min-{eps:g}"
        cap = obj_B0 * (1 + eps / 100) + (args.zero_tol * obj_B0 if eps == 0 else 0.0)
        print(f"\n{'=' * 80}\n[CO2-비용] {label}: 목적함수 ≤ {cap / 1e9:.4f}조 → 배출량 최소화\n{'=' * 80}", flush=True)
        t0 = time.time()

        def stage1(nn, sns):
            nn.model.add_constraints(budget_expr(nn, sns) * (1 / BUDGET_SCALE) <= cap / BUDGET_SCALE,
                                     name='CO2MGA-budget')
            fallback_cap(nn, sns, aux_energy)
            nn.model.add_objective(emission_expr(nn, sns) * (1 / E_SCALE), overwrite=True)
            print(f"  [1단계] 비용 상한 {cap / 1e9:.4f}조, 목적함수 → 배출량", flush=True)

        n = m.create_network(input_data)
        try:
            ok1 = solve(n, stage1)
        except Exception as e:
            print(f"[CO2-비용] {label} 1단계 오류: {e}", flush=True)
            ok1 = False
        if not ok1:
            rows.append({'case': label, 'eps_pct': eps, 'cost_cap': cap, 'ok': False, 'stage': 1}); flush()
            continue
        E1 = co2_emissions(n) * 1e6
        Z1 = system_cost(n)[1]
        s1 = {'stage1_CO2_Mt': E1 / 1e6, 'stage1_objective': Z1, 'stage1_C_sys': Z1 - penalty_value(n) - aux_capex(n)}
        print(f"[CO2-비용] {label} 1단계: CO2={E1 / 1e6:.3f}Mt (B0 {E_B0 / 1e6:.3f}), "
              f"목적함수={Z1 / 1e9:.4f}조", flush=True)
        del n

        def stage2(nn, sns):
            fallback_cap(nn, sns, aux_energy)
            nn.model.add_constraints(emission_expr(nn, sns) * (1 / E_SCALE) <= E1 * (1 + E_TOL) / E_SCALE,
                                     name='CO2MGA-emission')
            print(f"  [2단계] 배출량 ≤ {E1 * (1 + E_TOL) / 1e6:.3f}Mt, 목적함수 = 원래 비용", flush=True)

        n = m.create_network(input_data)
        try:
            ok2 = solve(n, stage2)
        except Exception as e:
            print(f"[CO2-비용] {label} 2단계 오류: {e}", flush=True)
            ok2 = False
        extra = dict(eps_pct=eps, cost_cap=cap, stage=2, **s1)
        if ok2:
            obj = float(n.objective)
            C2 = sys_cost(n)
            extra.update(C_sys=C2, objective_full=obj, penalty_cost=penalty_value(n), slack_capex=aux_capex(n),
                         objective_increase_pct=100 * (obj / obj_B0 - 1),
                         cost_increase_pct=100 * (C2 / C_B0 - 1), within_cap=bool(obj <= cap * (1 + 1e-6)),
                         C_sys_statistics=system_cost(n)[0] - aux_capex(n))
        rows.append(record(n, label, ok2, t0, extra))
        if ok2:
            regs += regional_capacity(n, label)
            n.export_to_netcdf(os.path.join(out_dir, f'{label}.nc'))
            r = rows[-1]
            print(f"[CO2-비용] {label} 완료 ({r['elapsed_min']}분): 목적함수 +{r['objective_increase_pct']:.3f}%, "
                  f"slack 제외 비용 {r['cost_increase_pct']:+.3f}% "
                  f"CO2={r['CO2_Mt']:.2f}Mt (B0 대비 {r['CO2_Mt'] - rows[0]['CO2_Mt']:+.2f})", flush=True)
        flush()
        del n

    df = pd.DataFrame(rows)
    pd.set_option('display.width', 250)
    cols = [c for c in ['case', 'ok', 'eps_pct', 'C_sys', 'cost_increase_pct', 'CO2_Mt', 'stage1_CO2_Mt',
                        'curtail_%', 'slack_TWh', 'heat_fallback_TWh'] if c in df]
    print(df[cols].round(4).to_string(index=False))
    print(f"\n요약 파일: {out_dir}")


if __name__ == '__main__':
    main()
