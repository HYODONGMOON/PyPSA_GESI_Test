"""
비용 상한 기반 최소 재생설비 대안 분석 (MGA, 2040 상한해제 기준)

interface/통합 입력 파일은 수정하지 않는다. PyPSA_GUI.optimize_network를 그대로 사용하고,
그 안의 extra_functionality(CO2 통합·DR·DC/Fab RE 제약) 뒤에 MGA 제약과 목적함수 교체만 덧붙인다.
(pypsa optimize_mga는 extra_functionality를 받지 않아 사용자 정의 제약이 빠지므로 쓰지 않는다)

공통 조건
  - 전력 slack(EL/EV slack·failsafe) 발전 0으로 고정 (B0 포함)
  - 시스템비용 C = PyPSA 목적함수 − 보조공급 페널티(열/수소 Fallback × 한계비용)
    · 목적함수 = 신규 설비 연간화 투자비 + 전체 운전비 (확장설비의 기존 용량 고정비는 상수로 제외)
  - MGA 단계에서는 Fallback 연간 공급량을 지역별 B0 수준 이하로 제한

단계
  B0      : 비용 최소화 → C_B0
  MGA-ε 1 : C ≤ (1+ε)·C_B0 (ε=0이면 수치 허용오차 1e-5) 하에서 Σ(PV·WT 신규 용량) 최소화 → K*
  MGA-ε 2 : Σ(PV·WT 신규 용량) ≤ K*(1+1e-4) 하에서 원래 목적함수(비용) 최소화
  점검    : 출력제한 페널티 30/100 (현재 입력, slack 0)

사용법:
  python run_mga_sweep.py [--eps 0 0.5 1 2 3 4 5] [--no-penalty]
"""
import argparse
import json
import os
import time
from datetime import datetime

import pandas as pd
import xarray as xr

import PyPSA_GUI as m
from run_demand_sweep import scale_demand
from run_ess_cost_sweep import summarize
from run_flex_sweep import apply_curtailment_penalty, penalty_in_objective
from run_regional_cost_sweep import apply_potential_cap, regional_capacity

K_TOL = 1e-4
COST_SCALE = 1e6
TOL_ENV = 'PYPSA_CPLEX_CONVERGETOL'


def load_input():
    m.ensure_integrated_input()
    d = m.read_input_data(m.INPUT_FILE)
    d = m.standardize_bus_names_in_input(d)
    if 'generators' in d and not d['generators'].empty:
        d['generators']['committable'] = False
    return d


def power_slack_gens(n):
    return [g for g in n.generators.index if 'slack' in g.lower() or 'failsafe' in g.lower()]


def fallback_gens(n):
    return [g for g in n.generators.index if 'fallback' in g.lower()]


def re_ext_gens(n):
    g = n.generators
    return list(g.index[g.carrier.isin(['solar', 'wind']) & g.p_nom_extendable.astype(bool)])


def fix_power_slack_zero(n):
    sl = power_slack_gens(n)
    n.generators.loc[sl, 'p_max_pu'] = 0.0
    n.generators.loc[sl, 'p_min_pu'] = 0.0
    cols = [c for c in sl if c in n.generators_t.p_max_pu.columns]
    if cols:
        n.generators_t.p_max_pu.drop(columns=cols, inplace=True)
    print(f"[전력 slack 0 고정] {len(sl)}개: {sl}", flush=True)


def penalty_value(n):
    fb = fallback_gens(n) + power_slack_gens(n)
    w = n.snapshot_weightings.objective
    mc = n.get_switchable_as_dense('Generator', 'marginal_cost')[fb]
    return float((n.generators_t.p[fb] * mc).mul(w, axis=0).sum().sum())


def co2_emissions(n):
    """CO2_limit(primary_energy)과 같은 산식: 발전기 p/eff × 배출계수.
    CHP 연료는 LNG_Fuel_Supply(gas) 발전기로 공급되므로 여기에 이미 포함된다."""
    w = n.snapshot_weightings.generators
    em = n.carriers['co2_emissions']
    em = em[em != 0]
    g = n.generators[n.generators.carrier.isin(em.index)]
    eff = n.get_switchable_as_dense('Generator', 'efficiency')[g.index]
    return float((n.generators_t.p[g.index] / eff * g.carrier.map(em)).mul(w, axis=0).sum().sum()) / 1e6


SOLVER_OVERRIDES = {}


class _OptimizeHook:
    """n.optimize 호출 시 기존 extra_functionality 다음에 추가 함수를 실행"""

    def __init__(self, acc, extra):
        self._acc, self._extra = acc, extra

    def __call__(self, *a, **kw):
        if SOLVER_OVERRIDES and kw.get('solver_options') is not None:
            kw['solver_options'] = {**kw['solver_options'], **SOLVER_OVERRIDES}
            print(f"  [MGA] 솔버 옵션 변경: {SOLVER_OVERRIDES}", flush=True)
        inner = kw.get('extra_functionality')

        def combined(n, sns):
            if inner:
                inner(n, sns)
            self._extra(n, sns)
        kw['extra_functionality'] = combined
        return self._acc(*a, **kw)

    def __getattr__(self, k):
        return getattr(self._acc, k)


def solution_assigned(n):
    """optimize_network는 'unknown' 상태도 수락하므로 해가 실제로 반영됐는지 확인"""
    re = re_ext_gens(n)
    g = n.generators.loc[re]
    return bool(g.p_nom_opt.notna().all() and (g.p_nom_opt >= g.p_nom - 1.0).all()
                and not n.generators_t.p.empty)


def solve(n, extra=None):
    if extra is not None:
        n.optimize = _OptimizeHook(n.optimize, extra)
    for tol in (None, '1e-5'):
        if tol:
            print(f"[재시도] convergetol={tol}", flush=True)
            os.environ[TOL_ENV] = tol
        try:
            ok = bool(m.optimize_network(n)) and solution_assigned(n)
        finally:
            os.environ.pop(TOL_ENV, None)
        if ok:
            return True
        print("[MGA] 해 미반영 또는 실패", flush=True)
    return False


def system_cost(n):
    """PyPSA 목적함수와 같은 정의(B0에서 일치 확인)로 사후 계산한 비용 − 보조공급 페널티"""
    st = n.statistics
    full = float(st.capex().sum() + st.opex().sum() - st.installed_capex().sum())
    return full - penalty_value(n), full


def cost_expr(n, sns):
    """원래 목적함수 식 − Fallback/slack 페널티 식 (시스템비용 식)"""
    mdl = n.model
    obj = mdl.objective
    obj = obj.expression if hasattr(obj, 'expression') else obj
    pen = fallback_gens(n) + power_slack_gens(n)
    w = n.snapshot_weightings.objective.loc[sns]
    mc = n.get_switchable_as_dense('Generator', 'marginal_cost', snapshots=sns)[pen].mul(w, axis=0)
    coef = xr.DataArray(mc.values, coords={'snapshot': sns, 'Generator': pen}, dims=['snapshot', 'Generator'])
    p = mdl.variables['Generator-p'].sel(Generator=pen, snapshot=sns)
    return obj - (p * coef).sum()


def fallback_cap(n, sns, fb_energy):
    mdl = n.model
    w = n.snapshot_weightings.generators.loc[sns]
    wd = xr.DataArray(w.values, coords={'snapshot': sns}, dims=['snapshot'])
    for g, e in fb_energy.items():
        p = mdl.variables['Generator-p'].sel(Generator=g, snapshot=sns)
        mdl.add_constraints((p * wd).sum() <= e * (1 + 1e-6) + 1.0, name=f'MGA-fallback-cap-{g}')


def re_new_expr(n):
    """PV·WT 총용량 변수 합 (기존 용량은 상수이므로 최소화 대상으로는 신규 용량과 동일)"""
    re = re_ext_gens(n)
    return n.model.variables['Generator-p_nom'].sel({'Generator-ext': re}).sum()


def re_new_mw(n):
    re = re_ext_gens(n)
    g = n.generators.loc[re]
    return float((g.p_nom_opt - g.p_nom).sum())


def record(n, label, ok, t0, extra):
    row = summarize(n, label, 1.0, ok, time.time() - t0)
    if ok:
        pen = penalty_value(n)
        row['penalty_cost'] = pen
        row['RE_GW'] = row['PV_GW'] + row['WT_GW']
        row['RE_added_GW'] = row['PV_added_GW'] + row['WT_added_GW']
        row['CO2_Mt'] = co2_emissions(n)
        st = n.stores[n.stores.index.str.endswith('_ESS')]
        row['ESS_GW_added'] = (n.links.loc[n.links.index.str.endswith('_ESS_Ch'), 'p_nom_opt']
                               - n.links.loc[n.links.index.str.endswith('_ESS_Ch'), 'p_nom']).sum() / 1e3
        row['ESS_GWh_existing'] = st.e_nom.sum() / 1e3
        h2 = n.stores[n.stores.index.str.endswith('_H2_Storage')]
        row['H2_storage_GWh'] = h2.e_nom_opt.sum() / 1e3
        hs = n.stores[n.stores.index.str.endswith('_H_Storage')]
        row['heat_storage_GWh'] = hs.e_nom_opt.sum() / 1e3
        evd = n.stores[n.stores.index.str.endswith('_EV_DSM')]
        row['EV_DSM_GWh'] = evd.e_nom_opt.sum() / 1e3
        el = n.links.index[n.links.index.str.endswith('_Electrolyser')]
        row['electrolyser_GW'] = n.links.loc[el, 'p_nom_opt'].sum() / 1e3
        nu = n.generators.index[n.generators.carrier == 'nuclear']
        row['nuc_TWh'] = n.generators_t.p[nu].sum().sum() / 1e6
        pv = n.generators.index[n.generators.carrier == 'solar']
        wt = n.generators.index[n.generators.carrier == 'wind']
        row['PV_TWh'] = n.generators_t.p[pv].sum().sum() / 1e6
        row['WT_TWh'] = n.generators_t.p[wt].sum().sum() / 1e6
    row.update(extra)
    return row


POTENTIAL_CAP = None
DEMAND_SCALE = {}


def build(input_data):
    n = m.create_network(input_data)
    fix_power_slack_zero(n)
    if POTENTIAL_CAP:
        apply_potential_cap(n, POTENTIAL_CAP)
    for kind, factor in DEMAND_SCALE.items():
        scale_demand(n, (kind,), factor)
    return n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--eps', type=float, nargs='+', default=[0, 0.5, 1, 2, 3, 4, 5])
    ap.add_argument('--no-penalty', action='store_true')
    ap.add_argument('--zero-tol', type=float, default=1e-5, help='MGA-0 비용 상한의 상대 허용오차')
    ap.add_argument('--b0-dir', help='B0_reference.json·mga_summary.csv가 있는 기존 결과 폴더 (B0 재사용)')
    ap.add_argument('--potential-cap', type=float, help='지역별 PV·WT 상한 = max(기술적 잠재량 × 비율, 기존 설비)')
    ap.add_argument('--barrier-alg', type=int, choices=[0, 1, 2, 3],
                    help='MGA 단계의 CPLEX barrier.algorithm (기본: PyPSA_GUI 설정 3)')
    ap.add_argument('--demand', nargs='+', default=[], metavar='KIND=FACTOR',
                    help='전력수요 배율, 예: DC=0.5 Fab=0.916 (KIND: EL, DC, Fab)')
    ap.add_argument('--tag', default='', help='결과 폴더 이름 접두어')
    args = ap.parse_args()

    global POTENTIAL_CAP
    POTENTIAL_CAP = args.potential_cap
    for item in args.demand:
        kind, factor = item.split('=')
        DEMAND_SCALE[kind] = float(factor)
    if args.barrier_alg is not None:
        SOLVER_OVERRIDES['barrier.algorithm'] = args.barrier_alg
    tag = (f'{args.tag}_' if args.tag else '') + (f"cap{int(round(args.potential_cap * 100))}_" if args.potential_cap else '')
    out_dir = os.path.join('results', f"mga_sweep_{tag}{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    os.makedirs(out_dir, exist_ok=True)
    rows, regs = [], []

    def flush():
        pd.DataFrame(rows).to_csv(os.path.join(out_dir, 'mga_summary.csv'), index=False, encoding='utf-8-sig')
        pd.DataFrame(regs).to_csv(os.path.join(out_dir, 'regional_capacity.csv'), index=False, encoding='utf-8-sig')

    input_data = load_input()

    if args.b0_dir:
        # ── B0 재사용 (같은 입력·같은 조건으로 계산된 결과) ─────────────────
        with open(os.path.join(args.b0_dir, 'B0_reference.json'), encoding='utf-8') as f:
            ref = json.load(f)
        C_B0, fb_energy, K_B0 = ref['C_B0'], ref['fallback_energy_MWh'], ref['RE_new_MW']
        prev = pd.read_csv(os.path.join(args.b0_dir, 'mga_summary.csv'))
        rows.append(prev[prev.case == 'B0'].iloc[0].to_dict())
        preg = pd.read_csv(os.path.join(args.b0_dir, 'regional_capacity.csv'))
        regs += preg[preg.case == 'B0'].to_dict('records')
        print(f"[MGA] B0 재사용: {args.b0_dir}, C_B0={C_B0 / 1e9:.4f}조, PV·WT 신규={K_B0 / 1e3:.2f}GW", flush=True)
    else:
        # ── B0 ────────────────────────────────────────────────────────────
        print(f"\n{'=' * 80}\n[MGA] B0: 비용 최소화 (전력 slack 0)\n{'=' * 80}", flush=True)
        t0 = time.time()
        n = build(input_data)
        re_list = re_ext_gens(n)
        print(f"[MGA] 신규용량 대상(PV·WT 확장가능) {len(re_list)}개: {re_list}", flush=True)
        if not solve(n):
            print("[MGA] B0 실패 — 중단")
            return
        obj_B0 = float(n.objective)
        pen_B0 = penalty_value(n)
        C_B0 = obj_B0 - pen_B0
        _, stat_cost = system_cost(n)
        fb_energy = {g: float(n.generators_t.p[g].mul(n.snapshot_weightings.generators).sum())
                     for g in fallback_gens(n)}
        K_B0 = re_new_mw(n)
        base = dict(case='B0', eps_pct=None, cost_cap=None, C_sys=C_B0, objective_full=obj_B0,
                    stage1_K_MW=K_B0, statistics_cost_check=stat_cost)
        rows.append(record(n, 'B0', True, t0, base))
        regs += regional_capacity(n, 'B0')
        n.export_to_netcdf(os.path.join(out_dir, 'B0.nc'))
        with open(os.path.join(out_dir, 'B0_reference.json'), 'w', encoding='utf-8') as f:
            json.dump({'objective': obj_B0, 'penalty': pen_B0, 'C_B0': C_B0, 'statistics_cost': stat_cost,
                       'RE_new_MW': K_B0, 'fallback_energy_MWh': fb_energy}, f, ensure_ascii=False, indent=1)
        print(f"[MGA] B0 목적함수={obj_B0 / 1e9:.4f}조, 페널티={pen_B0 / 1e9:.4f}조, C_B0={C_B0 / 1e9:.4f}조, "
              f"statistics 검산={stat_cost / 1e9:.4f}조, PV·WT 신규={K_B0 / 1e3:.2f}GW", flush=True)
        flush()
        del n

    # ── MGA-ε ─────────────────────────────────────────────────────────
    for eps in args.eps:
        label = f"MGA-{eps:g}"
        cap = C_B0 * (1 + eps / 100) + (args.zero_tol * C_B0 if eps == 0 else 0.0)
        print(f"\n{'=' * 80}\n[MGA] {label}: C ≤ {cap / 1e9:.4f}조 (B0 {C_B0 / 1e9:.4f}조 × {1 + eps / 100:.3f})"
              f" → PV·WT 신규용량 최소화\n{'=' * 80}", flush=True)
        t0 = time.time()
        def stage1(nn, sns):
            ce = cost_expr(nn, sns)
            nn.model.add_constraints(ce * (1 / COST_SCALE) <= cap / COST_SCALE, name='MGA-budget')
            fallback_cap(nn, sns, fb_energy)
            nn.model.add_objective(re_new_expr(nn), overwrite=True)
            print(f"  [MGA 1단계] 비용 상한 {cap / 1e9:.4f}조 추가, 목적함수 → Σ PV·WT p_nom", flush=True)

        n = build(input_data)
        try:
            ok1 = solve(n, stage1)
        except Exception as e:
            print(f"[MGA] {label} 1단계 오류: {e}", flush=True)
            ok1 = False
        if not ok1:
            rows.append({'case': label, 'eps_pct': eps, 'cost_cap': cap, 'ok': False, 'stage': 1}); flush()
            continue
        K = re_new_mw(n)
        C1, _ = system_cost(n)
        s1 = {'stage1_objective_MW': float(n.objective), 'stage1_K_MW': K, 'stage1_C_sys': C1,
              'stage1_PV_GW': n.generators.loc[n.generators.carrier == 'solar', 'p_nom_opt'].sum() / 1e3,
              'stage1_WT_GW': n.generators.loc[n.generators.carrier == 'wind', 'p_nom_opt'].sum() / 1e3}
        print(f"[MGA] {label} 1단계: K*={K / 1e3:.2f}GW (B0 {K_B0 / 1e3:.2f}GW), 비용={C1 / 1e9:.4f}조", flush=True)
        del n

        def stage2(nn, sns):
            fallback_cap(nn, sns, fb_energy)
            existing = float(nn.generators.loc[re_ext_gens(nn), 'p_nom'].sum())
            nn.model.add_constraints(re_new_expr(nn) <= existing + K * (1 + K_TOL), name='MGA-capacity')
            print(f"  [MGA 2단계] PV·WT 신규 ≤ {K * (1 + K_TOL) / 1e3:.3f}GW, 목적함수 = 원래 비용", flush=True)

        n = build(input_data)
        try:
            ok2 = solve(n, stage2)
        except Exception as e:
            print(f"[MGA] {label} 2단계 오류: {e}", flush=True)
            ok2 = False
        extra = dict(eps_pct=eps, cost_cap=cap, stage=2, **s1)
        if ok2:
            obj = float(n.objective)
            C2 = obj - penalty_value(n)
            extra.update(C_sys_statistics=system_cost(n)[0])
            extra.update(C_sys=C2, objective_full=obj, cost_increase_pct=100 * (C2 / C_B0 - 1),
                         within_cap=bool(C2 <= cap * (1 + 1e-7)))
        rows.append(record(n, label, ok2, t0, extra))
        if ok2:
            regs += regional_capacity(n, label)
            n.export_to_netcdf(os.path.join(out_dir, f'{label}.nc'))
            r = rows[-1]
            print(f"[MGA] {label} 완료 ({r['elapsed_min']}분): 비용 +{r['cost_increase_pct']:.3f}% "
                  f"PV={r['PV_GW']:.1f} WT={r['WT_GW']:.1f} RE={r['RE_GW']:.1f}GW "
                  f"(B0 대비 {r['RE_GW'] - rows[0]['RE_GW']:+.1f}) ESS={r['ESS_GW']:.1f}GW/{r['ESS_GWh']:.0f}GWh "
                  f"출력제한={r['curtail_%']:.2f}%", flush=True)
        flush()
        del n

    # ── 점검: 출력제한 페널티 30 / 100 (현재 입력) ──────────────────────
    if not args.no_penalty:
        for c in (30.0, 100.0):
            label = f"check_penalty_{int(c):03d}"
            print(f"\n{'=' * 80}\n[MGA 점검] {label}\n{'=' * 80}", flush=True)
            t0 = time.time()
            n = build(input_data)
            info = apply_curtailment_penalty(n, c)
            ok = solve(n)
            extra = {'curtail_penalty': c}
            if ok:
                obj = float(n.objective)
                C = obj - penalty_in_objective(n, c, *info) - penalty_value(n)
                extra.update(C_sys=C, objective_full=obj - penalty_in_objective(n, c, *info),
                             cost_increase_pct=100 * (C / C_B0 - 1), stage1_K_MW=re_new_mw(n))
            rows.append(record(n, label, ok, t0, extra))
            if ok:
                regs += regional_capacity(n, label)
            flush()
            del n

    df = pd.DataFrame(rows)
    pd.set_option('display.width', 250)
    cols = [c for c in ['case', 'ok', 'eps_pct', 'C_sys', 'cost_increase_pct', 'PV_GW', 'WT_GW', 'RE_GW',
                        'ESS_GW', 'ESS_GWh', 'curtail_%', 'CO2_Mt', 'slack_TWh', 'heat_fallback_TWh'] if c in df]
    print(df[cols].round(3).to_string(index=False))
    print(f"\n요약 파일: {out_dir}")


if __name__ == '__main__':
    main()
