"""
감축목표 달성에 필요한 추가 재생에너지(PV·풍력) 용량

기준 네트워크(B0.nc, 기존 분석 조건)에서 PV·풍력의 p_nom_max만 지역별 기술적 잠재량까지 풀고,
  배출량 ≤ 목표 (CO2_limit과 같은 발전기 기준 산식)
  보조공급(Fallback·slack) 연간량 ≤ 기준 + aux_cap_gwh (발전기별, 미공급으로 감축하지 못하게)
제약 아래 원래 목적함수(비용)를 최소화한다. 나머지 설비 조건은 기존과 같다.

  --mode opt : PV·풍력 모두 증설 가능 (최소비용 조합)
  --mode pv  : PV만 증설
  --mode wt  : 풍력만 증설

사용법:
  python run_re_need.py --tag 2035 --b0-nc <B0.nc> --target 88.3 --modes opt pv wt
"""
import argparse
import os
import time
from datetime import datetime

import pandas as pd
import pypsa
import xarray as xr

from run_co2_cost_sweep import aux_capex, aux_gens, emission_expr
from run_mga_sweep import co2_emissions, penalty_value, record, solve
from run_regional_cost_sweep import TECH_POTENTIAL_GW

E_SCALE = 1e3


def open_re(n, mode):
    g = n.generators
    for tech, car in (('PV', 'solar'), ('WT', 'wind')):
        for r, pot in TECH_POTENTIAL_GW[tech].items():
            nm = f'{r}_{tech}'
            if nm not in g.index:
                continue
            assert g.at[nm, 'carrier'] == car
            g.at[nm, 'p_nom_extendable'] = True
            g.at[nm, 'p_nom_min'] = g.at[nm, 'p_nom']
            if mode == 'opt' or mode == tech.lower():
                g.at[nm, 'p_nom_max'] = max(pot * 1e3, g.at[nm, 'p_nom'])
            else:
                g.at[nm, 'p_nom_max'] = g.at[nm, 'p_nom']


def re_table(n, label):
    g = n.generators
    rows = []
    for tech in ('PV', 'WT'):
        for r, pot in TECH_POTENTIAL_GW[tech].items():
            nm = f'{r}_{tech}'
            rows.append(dict(case=label, tech=tech, region=r, exist_GW=g.at[nm, 'p_nom'] / 1e3,
                             opt_GW=g.at[nm, 'p_nom_opt'] / 1e3, add_GW=(g.at[nm, 'p_nom_opt'] - g.at[nm, 'p_nom']) / 1e3,
                             potential_GW=pot, capital_cost=g.at[nm, 'capital_cost']))
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--tag', required=True)
    ap.add_argument('--b0-nc', required=True)
    ap.add_argument('--target', type=float, required=True, help='배출 목표 (Mt)')
    ap.add_argument('--modes', nargs='+', default=['opt', 'pv', 'wt'])
    ap.add_argument('--aux-cap-gwh', type=float, default=1.0)
    args = ap.parse_args()

    out_dir = os.path.join('results', f"re_need_{args.tag}_{datetime.now():%Y%m%d_%H%M%S}")
    os.makedirs(out_dir, exist_ok=True)
    b0 = pypsa.Network(args.b0_nc)
    Z_B0, E_B0 = float(b0.objective), co2_emissions(b0) * 1e6
    C_B0 = Z_B0 - penalty_value(b0) - aux_capex(b0)
    w = b0.snapshot_weightings.generators
    aux_B0 = {g: float(b0.generators_t.p[g].mul(w).sum()) for g in aux_gens(b0)}
    print(f"[재생필요] {args.tag} B0 Z={Z_B0 / 1e9:.5f}조, C={C_B0 / 1e9:.4f}조, E={E_B0 / 1e6:.3f}Mt, "
          f"목표 {args.target}Mt", flush=True)
    rows = [record(b0, 'B0', True, time.time(), dict(target_Mt=args.target, objective_full=Z_B0, C_sys=C_B0))]
    regs = re_table(b0, 'B0')
    del b0

    for mode in args.modes:
        label = f'RE-{mode}'
        print(f"\n{'=' * 80}\n[재생필요] {label}: 배출 ≤ {args.target}Mt\n{'=' * 80}", flush=True)
        t0 = time.time()

        def ef(nn, sns):
            mdl = nn.model
            mdl.add_constraints(emission_expr(nn, sns) * (1 / E_SCALE) <= args.target * 1e6 / E_SCALE,
                                name='RE-need-CO2')
            wd = xr.DataArray(nn.snapshot_weightings.generators.loc[sns].values, coords={'snapshot': sns},
                              dims=['snapshot'])
            for g, e in aux_B0.items():
                p = mdl.variables['Generator-p'].sel(Generator=g, snapshot=sns)
                mdl.add_constraints((p * wd).sum() <= e + args.aux_cap_gwh * 1e3, name=f'RE-need-aux-{g}')

        n = pypsa.Network(args.b0_nc)
        open_re(n, mode)
        try:
            ok = solve(n, ef)
        except Exception as e:
            print(f"[재생필요] {label} 오류: {e}", flush=True)
            ok = False
        if not ok:
            rows.append(dict(case=label, ok=False, target_Mt=args.target))
            pd.DataFrame(rows).to_csv(os.path.join(out_dir, 'summary.csv'), index=False, encoding='utf-8-sig')
            continue
        Z = float(n.objective)
        C = Z - penalty_value(n) - aux_capex(n)
        aux = sum(float(n.generators_t.p[g].mul(n.snapshot_weightings.generators).sum()) for g in aux_B0)
        rt = re_table(n, label)
        add_pv = sum(r['add_GW'] for r in rt if r['tech'] == 'PV')
        add_wt = sum(r['add_GW'] for r in rt if r['tech'] == 'WT')
        rows.append(record(n, label, True, t0, dict(
            target_Mt=args.target, objective_full=Z, C_sys=C, penalty_cost=penalty_value(n), slack_capex=aux_capex(n),
            objective_increase_pct=100 * (Z / Z_B0 - 1), cost_increase_pct=100 * (C / C_B0 - 1),
            aux_energy_diff_MWh=aux - sum(aux_B0.values()), PV_add_GW=add_pv, WT_add_GW=add_wt)))
        regs += rt
        n.export_to_netcdf(os.path.join(out_dir, f'{label}.nc'))
        r = rows[-1]
        print(f"[재생필요] {label}: PV +{add_pv:.1f}GW, 풍력 +{add_wt:.1f}GW, CO2 {r['CO2_Mt']:.2f}Mt, "
              f"slack 제외 비용 {r['cost_increase_pct']:+.2f}%, 보조공급 변화 {r['aux_energy_diff_MWh'] / 1e3:+.1f}GWh, "
              f"{r['elapsed_min']}분", flush=True)
        pd.DataFrame(rows).to_csv(os.path.join(out_dir, 'summary.csv'), index=False, encoding='utf-8-sig')
        pd.DataFrame(regs).to_csv(os.path.join(out_dir, 're_regional.csv'), index=False, encoding='utf-8-sig')
        del n

    print(f"\n요약 파일: {out_dir}")


if __name__ == '__main__':
    main()
