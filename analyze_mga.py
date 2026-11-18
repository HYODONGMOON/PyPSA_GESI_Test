"""
MGA 결과 분석: 저장된 네트워크(.nc)에서 전력·열·가스 수급, 비용 구성, 지역별 설비를 집계

사용법:
  python analyze_mga.py <결과폴더1> [<결과폴더2> ...]
  python analyze_mga.py <결과폴더 ...> --out <저장폴더>
  (B0.nc, MGA-*.nc, mga_summary.csv가 있는 폴더들을 합쳐 reports/mga_<날짜>/ 또는 --out 폴더에 저장)
"""
import glob
import os
import re
import sys
from datetime import datetime

import pandas as pd
import pypsa

from run_regional_cost_sweep import TECH_POTENTIAL_GW

OUT = os.path.join('reports', f"mga_{datetime.now().strftime('%Y%m%d')}")
REG = ['GGD', 'ICN', 'GWD', 'CBD', 'CND', 'JBD', 'GBD', 'BSN', 'GND', 'JND', 'JJD']


def kind(name):
    return re.sub(r'^[A-Z]{3}_', '', name.strip())


def analyze(n, case):
    w = n.snapshot_weightings.generators
    g, L, S = n.generators, n.links, n.stores
    P = n.generators_t.p.mul(w, axis=0).sum() / 1e6
    p0 = n.links_t.p0.mul(w, axis=0).sum() / 1e6
    p1 = n.links_t.p1.mul(w, axis=0).sum() / 1e6
    p2 = n.links_t.p2.mul(w, axis=0).sum() / 1e6 if 'p2' in n.links_t and not n.links_t.p2.empty else pd.Series(0.0, index=L.index)
    lk = L.index.map(kind)
    out = {'case': case}

    def lsum(s, *ks):
        return float(s[L.index[lk.isin(ks)]].sum())

    # 설비
    for car, key in (('solar', 'PV'), ('wind', 'WT')):
        idx = g.index[g.carrier == car]
        out[f'{key}_GW'] = g.loc[idx, 'p_nom_opt'].sum() / 1e3
        out[f'{key}_new_GW'] = (g.loc[idx, 'p_nom_opt'] - g.loc[idx, 'p_nom']).sum() / 1e3
        pot = n.generators_t.p_max_pu.reindex(columns=idx).fillna(g.loc[idx, 'p_max_pu']) * g.loc[idx, 'p_nom_opt']
        avail = pot.mul(w, axis=0).sum().sum() / 1e6
        out[f'{key}_TWh'] = float(P[idx].sum())
        out[f'{key}_curtail_TWh'] = avail - out[f'{key}_TWh']
        out[f'{key}_CF_%'] = 100 * out[f'{key}_TWh'] / (out[f'{key}_GW'] * w.sum() / 1e3)
    out['RE_GW'] = out['PV_GW'] + out['WT_GW']
    out['RE_new_GW'] = out['PV_new_GW'] + out['WT_new_GW']
    out['curtail_TWh'] = out['PV_curtail_TWh'] + out['WT_curtail_TWh']
    out['curtail_%'] = 100 * out['curtail_TWh'] / (out['curtail_TWh'] + out['PV_TWh'] + out['WT_TWh'])
    ess = S.index[S.index.map(kind) == 'ESS']
    out['ESS_GWh'] = S.loc[ess, 'e_nom_opt'].sum() / 1e3
    out['ESS_GW'] = L.loc[L.index[lk == 'ESS_Ch'], 'p_nom_opt'].sum() / 1e3
    out['ESS_new_GWh'] = (S.loc[ess, 'e_nom_opt'] - S.loc[ess, 'e_nom']).sum() / 1e3
    ess_re = S.index[S.index.map(kind) == 'ESS_RE']
    out['ESS_RE_GWh'] = S.loc[ess_re, 'e_nom_opt'].sum() / 1e3
    for k, key in (('Electrolyser', 'electrolyser_GW'), ('Fuelcell', 'fuelcell_GW')):
        out[key] = L.loc[L.index[lk == k], 'p_nom_opt'].sum() / 1e3
    out['H2_storage_GWh'] = S.loc[S.index[S.index.map(kind) == 'H2_Storage'], 'e_nom_opt'].sum() / 1e3

    # 전력 수급 (TWh)
    nu = g.index[g.carrier == 'nuclear']
    lng = [x for x in g.index[g.carrier == 'gas'] if x.endswith('_LNG')]
    fuel = [x for x in g.index[g.carrier == 'gas'] if 'Fuel_Supply' in x]
    out['nuclear_TWh'] = float(P[nu].sum())
    out['LNG_power_TWh'] = float(P[lng].sum())
    out['CHP_power_TWh'] = -lsum(p1, 'CHP')
    out['CHP_heat_TWh'] = -lsum(p2, 'CHP')
    out['CHP_fuel_TWh'] = lsum(p0, 'CHP')
    eff = n.get_switchable_as_dense('Generator', 'efficiency')
    out['LNG_power_fuel_TWh'] = float((n.generators_t.p[lng] / eff[lng]).mul(w, axis=0).sum().sum() / 1e6)
    out['gas_fuel_TWh'] = float(P[fuel].sum()) + out['LNG_power_fuel_TWh']
    out['HP_elec_TWh'] = lsum(p0, 'HP')
    out['HP_heat_TWh'] = -lsum(p1, 'HP')
    out['electrolyser_elec_TWh'] = lsum(p0, 'Electrolyser')
    out['fuelcell_power_TWh'] = -lsum(p1, 'Fuelcell')
    out['EV_charge_TWh'] = lsum(p0, 'EV_Charger', 'EV_Charger_V1G')
    out['V1G_charge_TWh'] = lsum(p0, 'EV_Charger_V1G')
    out['V2G_TWh'] = -lsum(p1, 'EV_V2G')
    out['ESS_charge_TWh'] = lsum(p0, 'ESS_Ch')
    out['ESS_loss_TWh'] = lsum(p0, 'ESS_Ch') + lsum(p1, 'ESS_Disch')
    hf = [x for x in g.index if 'H_Fallback' in x]
    out['heat_fallback_TWh'] = float(P.reindex(hf).fillna(0).sum())
    dr = g.index[g.carrier == 'DR']
    out['DR_TWh'] = float(P[dr].sum())
    load = n.loads_t.p.mul(w, axis=0).sum() / 1e6
    lkind = n.loads.index.map(kind)
    for k in ('Demand_EL', 'Demand_DC', 'Demand_Fab', 'Demand_H', 'Demand_EV', 'Demand_H2'):
        out[f'{k}_TWh'] = float(load[n.loads.index[lkind.isin([k, f'{k}_V1G'])]].sum())
    out['heat_storage_loss_TWh'] = lsum(p0, 'Heat_Ch') + lsum(p1, 'Heat_Disch')
    out['elec_demand_total_TWh'] = (out['Demand_EL_TWh'] + out['Demand_DC_TWh'] + out['Demand_Fab_TWh']
                                    + out['HP_elec_TWh'] + out['electrolyser_elec_TWh'] + out['EV_charge_TWh'])

    # 배출
    em = n.carriers['co2_emissions']
    em = em[em != 0]
    ge = g.index[g.carrier.isin(em.index)]
    out['CO2_Mt'] = float((n.generators_t.p[ge] / eff[ge] * g.loc[ge, 'carrier'].map(em)).mul(w, axis=0).sum().sum() / 1e6)
    gc = n.global_constraints
    out['CO2_price'] = float(-gc.at['CO2_limit', 'mu']) if 'CO2_limit' in gc.index else float('nan')
    el = [b for b in n.buses.index if b.endswith('_EL')]
    mp = n.buses_t.marginal_price[el]
    ld = n.loads_t.p[[c for c in n.loads.index if n.loads.at[c, 'bus'] in el]]
    lb = ld.T.groupby(n.loads.loc[ld.columns, 'bus']).sum().T.reindex(columns=el).fillna(0)
    out['EL_price_loadw'] = float((mp * lb).sum().sum() / lb.sum().sum())

    # 비용 구성 (조원): 신규 설비 연간화 투자비 + 운전비
    cost = {}
    for c, attr in (('Generator', 'p_nom'), ('Link', 'p_nom'), ('Store', 'e_nom'), ('Line', 's_nom')):
        df = n.static(c)
        ext = df.index[df[f'{attr}_extendable'].astype(bool)]
        if len(ext):
            capex = df.loc[ext, 'capital_cost'] * (df.loc[ext, f'{attr}_opt'] - df.loc[ext, attr])
            for nm, v in capex.items():
                key = {'solar': 'capex_PV', 'wind': 'capex_WT'}.get(df.at[nm, 'carrier'], f'capex_{kind(nm)}') \
                    if c == 'Generator' else f'capex_{kind(nm)}'
                cost[key] = cost.get(key, 0.0) + float(v)
    mcg = n.get_switchable_as_dense('Generator', 'marginal_cost')
    opg = (n.generators_t.p * mcg).mul(n.snapshot_weightings.objective, axis=0).sum()
    for nm, v in opg.items():
        car = g.at[nm, 'carrier']
        key = ('opex_penalty' if ('fallback' in nm.lower() or 'slack' in nm.lower())
               else 'opex_gas' if car == 'gas' else f'opex_{car}')
        cost[key] = cost.get(key, 0.0) + float(v)
    mcl = n.get_switchable_as_dense('Link', 'marginal_cost')
    opl = (n.links_t.p0 * mcl).mul(n.snapshot_weightings.objective, axis=0).sum()
    for nm, v in opl.items():
        cost[f'opex_{kind(nm)}'] = cost.get(f'opex_{kind(nm)}', 0.0) + float(v)
    if not S.empty and 'marginal_cost' in S:
        ops = (n.stores_t.p * n.get_switchable_as_dense('Store', 'marginal_cost')).mul(
            n.snapshot_weightings.objective, axis=0).sum()
        for nm, v in ops.items():
            cost[f'opex_store_{kind(nm)}'] = cost.get(f'opex_store_{kind(nm)}', 0.0) + float(v)
    cost = {k: v / 1e9 for k, v in cost.items() if abs(v) > 1e3}
    out['cost_total_check'] = sum(cost.values())
    out['cost_sys'] = out['cost_total_check'] - cost.get('opex_penalty', 0.0)
    out.update({f'c_{k}': v for k, v in cost.items()})

    # 지역별
    reg = []
    for r in REG:
        row = {'case': case, 'region': r}
        for tech in ('PV', 'WT'):
            nm = f'{r}_{tech}'
            row[f'{tech}_GW'] = g.at[nm, 'p_nom_opt'] / 1e3
            row[f'{tech}_exist_GW'] = g.at[nm, 'p_nom'] / 1e3
            row[f'{tech}_pot_%'] = row[f'{tech}_GW'] / TECH_POTENTIAL_GW[tech][r] * 100
        hp = f'{r}_HP'
        row['HP_elec_TWh'] = float(p0.get(hp, 0.0))
        row['CHP_heat_TWh'] = float(-p2.get(f'{r}_CHP', 0.0))
        reg.append(row)

    # 송전 혼잡
    lines = n.lines.index
    s = n.lines_t.p0.abs()
    smax = n.lines.s_nom_opt * n.get_switchable_as_dense('Line', 's_max_pu')
    out['line_congestion_h'] = float((s >= 0.99 * smax).sum().sum()) if len(lines) else 0.0
    hv = L.index[lk == 'HVDC']
    out['HVDC_TWh'] = float(n.links_t.p0[hv].abs().mul(w, axis=0).sum().sum() / 1e6)
    return out, reg


def main(dirs):
    os.makedirs(OUT, exist_ok=True)
    summ = pd.concat([pd.read_csv(os.path.join(d, 'mga_summary.csv')) for d in dirs], ignore_index=True)
    summ = summ[summ.ok.astype(bool)].drop_duplicates('case', keep='last')
    nat, regs = [], []
    files = {}
    for d in dirs:
        for f in glob.glob(os.path.join(d, '*.nc')):
            files[os.path.splitext(os.path.basename(f))[0]] = f
    order = ['B0'] + sorted([c for c in files if c != 'B0'], key=lambda c: float(c.split('-', 1)[1]))
    for c in order:
        print(f'[분석] {c}', flush=True)
        n = pypsa.Network(files[c])
        o, r = analyze(n, c)
        nat.append(o); regs += r
    nat = pd.DataFrame(nat).set_index('case')
    keep = ['eps_pct', 'cost_cap', 'C_sys', 'cost_increase_pct', 'stage1_K_MW', 'stage1_C_sys', 'elapsed_min']
    nat = nat.join(summ.set_index('case')[[k for k in keep if k in summ]])
    nat.to_csv(os.path.join(OUT, 'mga_national.csv'), encoding='utf-8-sig')
    pd.DataFrame(regs).to_csv(os.path.join(OUT, 'mga_regional.csv'), index=False, encoding='utf-8-sig')
    summ.to_csv(os.path.join(OUT, 'mga_summary_all.csv'), index=False, encoding='utf-8-sig')
    pd.set_option('display.width', 250); pd.set_option('display.max_rows', 300)
    print(nat.T.round(3).to_string())
    print(f'저장: {OUT}')


if __name__ == '__main__':
    args = sys.argv[1:]
    if '--out' in args:
        i = args.index('--out')
        OUT = args[i + 1]
        args = args[:i] + args[i + 2:]
    main(args)
