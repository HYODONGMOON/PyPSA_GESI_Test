"""
비용 상한별 최소 배출 대안 — 탄소가격(λ) 탐색 방식

비용 상한 제약 + 배출량 목적함수(run_co2_cost_sweep)는 CPLEX Barrier가 수렴하지 못해서,
같은 효율 경계(LP Pareto front)를 가중합 목적함수로 구한다.

  목적함수 = PyPSA 목적함수 Z (기존 분석과 같은 비용 정의, slack 포함) + λ·E
  비용 상한 Z ≤ (1+ε)·Z_B0 를 만족하는 가장 큰 λ의 해 = 해당 상한에서 배출량 최소 해
  ε=0 : λ=1천원/t (비용 증가 ≤ λ·ΔE ≈ 수억원 → 최소비용 해 중 배출량 최저)
  모델 비용 단위는 천원 → λ 단위는 천원/t (열 이름 lambda_won_per_t 도 천원/t)

  - Z, E 는 λ항을 뺀 사후 계산값. 보고용 시스템비용은 slack 비용을 후처리에서 제외
  - Fallback·slack 공급량 상한은 두지 않는다 (B0 수준 상한은 Barrier 수렴 실패 원인).
    페널티 단가(≈500만원/MWh)가 탐색 범위의 λ보다 훨씬 커서 보조공급이 배출 감축 수단이 되지 않는다
  - interface/통합 입력 파일은 수정하지 않는다

사용법:
  python run_co2_price_search.py [--reuse-b0 <B0.nc 폴더>] [--ref-objective <원>] [--tag 2035]
                                 [--eps 0 1 3 5 10] [--tol 0.1]
"""
import argparse
import json
import os
import shutil
import time
from datetime import datetime

import pandas as pd
import pypsa

import PyPSA_GUI as m
from run_co2_cost_sweep import aux_capex, aux_gens, emission_expr
from run_mga_sweep import co2_emissions, load_input, penalty_value, record, solve, system_cost
from run_regional_cost_sweep import regional_capacity

GRID = [30, 60, 80, 100, 130, 500]
LAM_MAX = 1e6


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--eps', type=float, nargs='+', default=[0, 1, 3, 5, 10])
    ap.add_argument('--reuse-b0', help='B0.nc가 있는 폴더 (없으면 현재 interface로 B0를 새로 계산)')
    ap.add_argument('--ref-objective', type=float, help='기준 결과(GUI) 목적함수, B0 일치 확인용')
    ap.add_argument('--tol', type=float, default=0.1, help='목표 비용증가율 허용 오차 (%p, 목표 이하 쪽)')
    ap.add_argument('--max-iter', type=int, default=6)
    ap.add_argument('--bisect', action='store_true', help='보간 대신 구간 중간값으로 λ 탐색')
    ap.add_argument('--resume', help='이전 실행 폴더 (lambda_log.csv의 해를 재사용)')
    ap.add_argument('--tag', default='2040')
    args = ap.parse_args()

    out_dir = args.resume or os.path.join('results', f"co2_price_search_{args.tag}_"
                                                     f"{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    os.makedirs(os.path.join(out_dir, 'lambda'), exist_ok=True)
    input_data = load_input()

    b0_path = os.path.join(out_dir, 'B0.nc')
    if args.reuse_b0:
        if os.path.abspath(os.path.join(args.reuse_b0, 'B0.nc')) != os.path.abspath(b0_path):
            shutil.copy(os.path.join(args.reuse_b0, 'B0.nc'), b0_path)
        n = pypsa.Network(b0_path)
    else:
        print(f"[탄소가격] B0: 비용 최소화 (현재 interface 조건)", flush=True)
        n = m.create_network(input_data)
        if not solve(n):
            print("[탄소가격] B0 실패 — 중단", flush=True)
            return
        n.export_to_netcdf(b0_path)
    Z_B0, E_B0 = float(n.objective), co2_emissions(n) * 1e6
    if args.ref_objective:
        print(f"[탄소가격] B0 목적함수 {Z_B0 / 1e9:.5f}조 / 기준 결과 {args.ref_objective / 1e9:.5f}조 "
              f"(차이 {100 * (Z_B0 / args.ref_objective - 1):+.4f}%)", flush=True)
    C_B0 = Z_B0 - penalty_value(n) - aux_capex(n)
    w = n.snapshot_weightings.generators
    aux_energy = {g: float(n.generators_t.p[g].mul(w).sum()) for g in aux_gens(n)}
    print(f"[탄소가격] B0 Z={Z_B0 / 1e9:.5f}조, slack 제외 C={C_B0 / 1e9:.4f}조, E={E_B0 / 1e6:.3f}Mt", flush=True)
    rows = [record(n, 'B0', True, time.time(), dict(eps_pct=0.0, lambda_won_per_t=0.0, objective_full=Z_B0,
                                                    C_sys=C_B0, penalty_cost=penalty_value(n),
                                                    slack_capex=aux_capex(n), objective_increase_pct=0.0,
                                                    cost_increase_pct=0.0))]
    regs = regional_capacity(n, 'B0')
    del n

    evals = {0.0: dict(Z=Z_B0, E=E_B0, path=os.path.join(out_dir, 'B0.nc'))}
    log = []
    if args.resume and os.path.exists(os.path.join(out_dir, 'lambda_log.csv')):
        log = pd.read_csv(os.path.join(out_dir, 'lambda_log.csv')).to_dict('records')
        for r in log:
            lam = float(r['lambda_won_per_t'])
            if r['ok']:
                evals[lam] = dict(Z=r['Z'], E=r['E_Mt'] * 1e6, path=os.path.join(out_dir, 'lambda', f'lam_{lam:g}.nc'))
        print(f"[탄소가격] 이전 해 {len(evals) - 1}개 재사용", flush=True)

    def run(lam):
        if lam in evals:
            return evals[lam]
        print(f"\n{'=' * 80}\n[탄소가격] λ={lam:,.2f}천원/t\n{'=' * 80}", flush=True)
        t0 = time.time()

        def ef(nn, sns):
            obj = nn.model.objective
            obj = obj.expression if hasattr(obj, 'expression') else obj
            nn.model.add_objective(obj + emission_expr(nn, sns) * lam, overwrite=True)

        nn = m.create_network(input_data)
        try:
            ok = solve(nn, ef)
        except Exception as e:
            print(f"[탄소가격] λ={lam:,.0f} 오류: {e}", flush=True)
            ok = False
        if not ok:
            evals[lam] = None
            log.append(dict(lambda_won_per_t=lam, ok=False))
            return None
        E = co2_emissions(nn) * 1e6
        Z = float(nn.objective) - lam * E
        nn.objective = Z
        path = os.path.join(out_dir, 'lambda', f'lam_{lam:g}.nc')
        nn.export_to_netcdf(path)
        r = dict(Z=Z, E=E, path=path, stat_check=system_cost(nn)[1], minutes=(time.time() - t0) / 60)
        evals[lam] = r
        log.append(dict(lambda_won_per_t=lam, ok=True, Z=Z, E_Mt=E / 1e6, dZ_pct=100 * (Z / Z_B0 - 1),
                        dE_Mt=(E - E_B0) / 1e6, statistics_Z=r['stat_check'], minutes=r['minutes']))
        pd.DataFrame(log).to_csv(os.path.join(out_dir, 'lambda_log.csv'), index=False, encoding='utf-8-sig')
        print(f"[탄소가격] λ={lam:,.2f}천원/t → 목적함수 {100 * (Z / Z_B0 - 1):+.3f}%, "
              f"CO2 {E / 1e6:.2f}Mt ({(E - E_B0) / 1e6:+.2f}), {r['minutes']:.1f}분", flush=True)
        del nn
        return r

    targets = sorted(e for e in args.eps if e > 0)
    if 0 in args.eps:
        run(1.0)
    for lam in GRID:
        run(lam)
    lam = GRID[-1]
    while targets and evals[lam] and evals[lam]['Z'] < Z_B0 * (1 + targets[-1] / 100) and lam < LAM_MAX:
        lam *= 2.5
        run(lam)

    def ok_items():
        return sorted((l, r) for l, r in evals.items() if r)

    def bracket(cap):
        lo = max(((l, r) for l, r in ok_items() if r['Z'] <= cap), key=lambda x: x[0])
        hi = [(l, r) for l, r in ok_items() if r['Z'] > cap and l > lo[0]]
        return lo, (min(hi, key=lambda x: x[0]) if hi else None)

    for t in targets:
        cap = Z_B0 * (1 + t / 100)
        for _ in range(args.max_iter):
            (l0, r0), hi = bracket(cap)
            if hi is None or 100 * (cap - r0['Z']) / Z_B0 <= args.tol:
                break
            l1, r1 = hi
            if hi[0] / max(l0, 1.0) < 1.0005:
                break
            frac = (cap - r0['Z']) / (r1['Z'] - r0['Z'])
            if args.bisect:
                frac = 0.5
            lam = l0 + frac * (l1 - l0) if l0 > 0 else l1 * max(frac, 0.1)
            lam = float(round(min(max(lam, l0 + 0.05 * (l1 - l0)), l1 - 0.05 * (l1 - l0)), 2))
            if run(lam) is None:
                evals.pop(lam, None)
                break

    for eps in args.eps:
        label = f"CO2min-{eps:g}"
        if eps == 0:
            lam, r = 1.0, evals.get(1.0)
        else:
            (lam, r), _ = bracket(Z_B0 * (1 + eps / 100))
        if not r:
            rows.append(dict(case=label, eps_pct=eps, ok=False))
            continue
        dst = os.path.join(out_dir, f'{label}.nc')
        shutil.copy(r['path'], dst)
        n = pypsa.Network(dst)
        C = r['Z'] - penalty_value(n) - aux_capex(n)
        rows.append(record(n, label, True, time.time(), dict(
            eps_pct=eps, lambda_won_per_t=lam, cost_cap=Z_B0 * (1 + eps / 100), objective_full=r['Z'],
            objective_increase_pct=100 * (r['Z'] / Z_B0 - 1), C_sys=C, penalty_cost=penalty_value(n),
            slack_capex=aux_capex(n), cost_increase_pct=100 * (C / C_B0 - 1),
            within_cap=bool(r['Z'] <= Z_B0 * (1 + max(eps, 1e-3) / 100)))))
        regs += regional_capacity(n, label)
        del n

    pd.DataFrame(rows).to_csv(os.path.join(out_dir, 'summary.csv'), index=False, encoding='utf-8-sig')
    pd.DataFrame(regs).to_csv(os.path.join(out_dir, 'regional_capacity.csv'), index=False, encoding='utf-8-sig')
    with open(os.path.join(out_dir, 'B0_reference.json'), 'w', encoding='utf-8') as f:
        json.dump({'objective': Z_B0, 'C_B0': C_B0, 'E_B0_t': E_B0, 'aux_energy_MWh': aux_energy}, f,
                  ensure_ascii=False, indent=1)
    df = pd.DataFrame(rows)
    cols = [c for c in ['case', 'ok', 'eps_pct', 'lambda_won_per_t', 'objective_increase_pct', 'cost_increase_pct',
                        'CO2_Mt'] if c in df]
    pd.set_option('display.width', 250)
    print(df[cols].round(4).to_string(index=False))
    print(f"\n요약 파일: {out_dir}")


if __name__ == '__main__':
    main()
