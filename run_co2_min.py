"""
비용 상한 없는 배출량 최소화 (기존 분석 조건, 설비 고정)

  목적함수 = PyPSA 목적함수 Z + λ·E, λ를 크게 두어 실제 설비로 줄일 수 있는 배출을 모두 줄인다.
  λ는 전력 slack(500만원/MWh)이 석탄(≈0.84 tCO2/MWh)을 대체하는 문턱(≈6,000천원/t)보다 낮게 둔다.
  모델 비용 단위는 천원 → λ 단위는 천원/t.

  네트워크는 현재 interface로 만들거나(--from-interface), 이전 B0.nc에서 다시 최적화한다(--b0-nc).

사용법:
  python run_co2_min.py --tag 2035 --from-interface --b0-nc <2035 B0.nc> [--lams 1000 3000]
  python run_co2_min.py --tag 2040 --b0-nc <2040 B0.nc> --check-b0 [--lams 1000 3000]
"""
import argparse
import os
import time
from datetime import datetime

import pandas as pd
import pypsa
import xarray as xr

import PyPSA_GUI as m
from run_co2_cost_sweep import aux_capex, aux_gens, emission_expr
from run_mga_sweep import co2_emissions, load_input, penalty_value, record, solve


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--tag', required=True)
    ap.add_argument('--b0-nc', required=True, help='기준(비용최소) 결과 nc')
    ap.add_argument('--from-interface', action='store_true', help='현재 interface로 네트워크 생성')
    ap.add_argument('--check-b0', action='store_true', help='λ=0으로 다시 풀어 B0 재현 확인')
    ap.add_argument('--lams', type=float, nargs='+', default=[1000, 3000])
    ap.add_argument('--aux-cap-gwh', type=float, help='보조공급 연간량 ≤ 기준 + 이 값(GWh, 발전기별)')
    args = ap.parse_args()

    out_dir = os.path.join('results', f"co2_min_{args.tag}_{datetime.now():%Y%m%d_%H%M%S}")
    os.makedirs(out_dir, exist_ok=True)
    input_data = load_input() if args.from_interface else None

    b0 = pypsa.Network(args.b0_nc)
    Z_B0, E_B0 = float(b0.objective), co2_emissions(b0) * 1e6
    C_B0 = Z_B0 - penalty_value(b0) - aux_capex(b0)
    w = b0.snapshot_weightings.generators
    aux_B0 = {g: float(b0.generators_t.p[g].mul(w).sum()) for g in aux_gens(b0)}
    print(f"[CO2최소] B0 Z={Z_B0 / 1e9:.5f}조, C={C_B0 / 1e9:.4f}조, E={E_B0 / 1e6:.3f}Mt", flush=True)
    rows = [record(b0, 'B0', True, time.time(), dict(lambda_kwon_t=0.0, objective_full=Z_B0, C_sys=C_B0,
                                                     objective_increase_pct=0.0, cost_increase_pct=0.0))]

    def build():
        if args.from_interface:
            return m.create_network(input_data)
        return pypsa.Network(args.b0_nc)

    cases = ([0.0] if args.check_b0 else []) + list(args.lams)
    for lam in cases:
        label = 'B0-check' if lam == 0 else f"CO2min-lam{lam:g}{'-auxcap' if args.aux_cap_gwh is not None else ''}"
        print(f"\n{'=' * 80}\n[CO2최소] {label}: λ={lam:,.0f}천원/t\n{'=' * 80}", flush=True)
        t0 = time.time()

        def ef(nn, sns, lam=lam):
            if args.aux_cap_gwh is not None and lam:
                mdl = nn.model
                wd = xr.DataArray(nn.snapshot_weightings.generators.loc[sns].values, coords={'snapshot': sns},
                                  dims=['snapshot'])
                for g, e in aux_B0.items():
                    p = mdl.variables['Generator-p'].sel(Generator=g, snapshot=sns)
                    mdl.add_constraints((p * wd).sum() <= e + args.aux_cap_gwh * 1e3, name=f'CO2MIN-aux-cap-{g}')
            if lam:
                obj = nn.model.objective
                obj = obj.expression if hasattr(obj, 'expression') else obj
                nn.model.add_objective(obj + emission_expr(nn, sns) * lam, overwrite=True)

        n = build()
        try:
            ok = solve(n, ef)
        except Exception as e:
            print(f"[CO2최소] {label} 오류: {e}", flush=True)
            ok = False
        if not ok:
            rows.append(dict(case=label, ok=False, lambda_kwon_t=lam))
            continue
        E = co2_emissions(n) * 1e6
        Z = float(n.objective) - lam * E
        n.objective = Z
        C = Z - penalty_value(n) - aux_capex(n)
        aux = {g: float(n.generators_t.p[g].mul(n.snapshot_weightings.generators).sum()) for g in aux_gens(n)}
        aux_diff = sum(aux.values()) - sum(aux_B0.values())
        rows.append(record(n, label, True, t0, dict(
            lambda_kwon_t=lam, objective_full=Z, C_sys=C, penalty_cost=penalty_value(n), slack_capex=aux_capex(n),
            objective_increase_pct=100 * (Z / Z_B0 - 1), cost_increase_pct=100 * (C / C_B0 - 1),
            aux_energy_diff_MWh=aux_diff)))
        n.export_to_netcdf(os.path.join(out_dir, f'{label}.nc'))
        print(f"[CO2최소] {label}: CO2 {E / 1e6:.2f}Mt ({(E - E_B0) / 1e6:+.2f}), 목적함수 {100 * (Z / Z_B0 - 1):+.3f}%, "
              f"slack 제외 비용 {100 * (C / C_B0 - 1):+.3f}%, 보조공급 변화 {aux_diff / 1e3:+.1f}GWh", flush=True)
        pd.DataFrame(rows).to_csv(os.path.join(out_dir, 'summary.csv'), index=False, encoding='utf-8-sig')
        del n

    pd.DataFrame(rows).to_csv(os.path.join(out_dir, 'summary.csv'), index=False, encoding='utf-8-sig')
    print(f"\n요약 파일: {out_dir}")


if __name__ == '__main__':
    main()
