"""
2035년 민감도·MGA 결과 집계 (analyze_mga.analyze 지표 + 석탄·수요 지표)

사용법:
  python analyze_2035.py <민감도 결과폴더> <MGA 결과폴더 ...> --out reports/sens2035_<날짜>
"""
import argparse
import glob
import os

import pandas as pd
import pypsa

from analyze_mga import analyze
from run_demand_sweep import el_demand_twh


def extra(n):
    g = n.generators
    w = n.snapshot_weightings.generators
    coal = g.index[g.carrier == 'coal']
    out = {'coal_GW': g.loc[coal, 'p_nom_opt'].sum() / 1e3,
           'coal_TWh': float(n.generators_t.p[coal].mul(w, axis=0).sum().sum() / 1e6),
           'CO2_cap_Mt': float(n.global_constraints.at['CO2_limit', 'constant']) / 1e6
           if 'CO2_limit' in n.global_constraints.index else float('nan'),
           'objective_full': float(n.objective)}
    out.update(el_demand_twh(n))
    out['demand_power_TWh'] = out['demand_EL_TWh'] + out['demand_DC_TWh'] + out['demand_Fab_TWh']
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('sens_dir')
    ap.add_argument('mga_dirs', nargs='+')
    ap.add_argument('--out', required=True)
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)

    files = {}
    for d in [args.sens_dir] + args.mga_dirs:
        files.update({os.path.splitext(os.path.basename(f))[0]: f for f in glob.glob(os.path.join(d, '*.nc'))})
    order = ['base', 'cap20', 'EL_m05', 'EL_m10', 'DC_m25', 'Fab_m25', 'DCFab_m25', 'DCFab_m50', 'A', 'B0', 'MGA-1', 'MGA-2']
    nat, regs = [], []
    for c in [c for c in order if c in files]:
        print(f'[분석] {c}', flush=True)
        n = pypsa.Network(files[c])
        o, r = analyze(n, c)
        o.update(extra(n))
        nat.append(o); regs += r
    nat = pd.DataFrame(nat).set_index('case')
    ms = pd.concat([pd.read_csv(os.path.join(d, 'mga_summary.csv')) for d in args.mga_dirs])
    ms = ms[ms.ok.astype(bool)].drop_duplicates('case', keep='last').set_index('case')
    keep = [k for k in ('eps_pct', 'cost_cap', 'cost_increase_pct', 'stage1_K_MW') if k in ms]
    nat = nat.join(ms[keep])
    nat.to_csv(os.path.join(args.out, 'national.csv'), encoding='utf-8-sig')
    pd.DataFrame(regs).to_csv(os.path.join(args.out, 'regional.csv'), index=False, encoding='utf-8-sig')
    pd.read_csv(os.path.join(args.sens_dir, 'sweep_summary.csv')).to_csv(
        os.path.join(args.out, 'sweep_summary.csv'), index=False, encoding='utf-8-sig')
    pd.set_option('display.width', 250); pd.set_option('display.max_rows', 300)
    print(nat.T.round(3).to_string())
    print(f'저장: {args.out}')


if __name__ == '__main__':
    main()
