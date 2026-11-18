"""
입력시트 B48(히트펌프)·B49(수전해) 연간 전력소비 목표 시험 — 2040 대표경로 ESS −30% 조건
interface.xlsx는 수정하지 않고 B48·B49 값만 메모리에서 지정한다.

사용법: python test_sector_targets_2040.py [--hp 27.6] [--ely 33.5]
"""
import argparse
import os
import tempfile

import pandas as pd

import PyPSA_GUI as m
import run_mga_sweep as r
import run_mga_2040_ess30 as e30


def chp_heat(n, w):
    p2 = n.links_t.get('p2', pd.DataFrame())
    cols = [l for l in p2.columns if n.links.at[l, 'bus2'].endswith('_H')]
    return float((-p2[cols]).mul(w, axis=0).sum().sum()) / 1e6


def summary(n, tag):
    w = n.snapshot_weightings.generators
    def link_in(sfx):
        idx = n.links.index[n.links.index.str.endswith(sfx)]
        return float(n.links_t.p0[idx].mul(w, axis=0).sum().sum()) / 1e6
    def gen(sfx):
        idx = n.generators.index[n.generators.index.str.endswith(sfx)]
        return float(n.generators_t.p[idx].mul(w, axis=0).sum().sum()) / 1e6
    def load(sfx):
        idx = n.loads.index[n.loads.index.str.strip().str.endswith(sfx)]
        return float(n.loads_t.p_set[idx].mul(w, axis=0).sum().sum()) / 1e6
    row = {
        'case': tag,
        'objective_조': n.objective / 1e9,
        'Demand_EL_TWh': load('_Demand_EL'),
        'Demand_H_TWh': load('_Demand_H'),
        'HP_elec_TWh': link_in('_HP'),
        'Electrolyser_elec_TWh': link_in('_Electrolyser'),
        'Fuelcell_H2_TWh': link_in('_Fuelcell'),
        'H2_offtake_TWh': gen('_H2_Offtake'),
        'Heat_fallback_TWh': gen('_H_Fallback_Gen'),
        'CHP_heat_TWh': chp_heat(n, w),
        'penalty_조': r.penalty_value(n) / 1e9,
        'PV_GW': n.generators.loc[n.generators.carrier == 'solar', 'p_nom_opt'].sum() / 1e3,
        'WT_GW': n.generators.loc[n.generators.carrier == 'wind', 'p_nom_opt'].sum() / 1e3,
        'Electrolyser_GW': n.links.loc[n.links.index.str.endswith('_Electrolyser'), 'p_nom_opt'].sum() / 1e3,
        'H2_storage_GWh': n.stores.loc[n.stores.index.str.endswith('_H2_Storage'), 'e_nom_opt'].sum() / 1e3,
        'CO2_Mt': r.co2_emissions(n),
    }
    reg = pd.DataFrame({
        'HP_TWh': {l.split('_')[0]: float(n.links_t.p0[l].mul(w).sum()) / 1e6
                   for l in n.links.index if l.endswith('_HP')},
        'Electrolyser_TWh': {l.split('_')[0]: float(n.links_t.p0[l].mul(w).sum()) / 1e6
                             for l in n.links.index if l.endswith('_Electrolyser')},
        'H2_offtake_TWh': {g.split('_')[0]: float(n.generators_t.p[g].mul(w).sum()) / 1e6
                           for g in n.generators.index if g.endswith('_H2_Offtake')},
        'Heat_fallback_TWh': {g.split('_')[0]: float(n.generators_t.p[g].mul(w).sum()) / 1e6
                              for g in n.generators.index if g.endswith('_H_Fallback_Gen')},
    })
    return row, reg


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--hp', type=float, default=27.6)
    ap.add_argument('--ely', type=float, default=33.5)
    args = ap.parse_args()

    m._read_sector_power_targets = lambda path: {'hp_TWh': args.hp, 'ely_TWh': args.ely}
    ref, _ = e30.load_reference()
    ess_cost = e30.ess_cost_from(ref)
    n = r.build(r.load_input())
    n.stores.loc[ess_cost.index, 'capital_cost'] = ess_cost.values
    print(f"[시험] ESS 단가 {ess_cost.iloc[0]:,.2f}, HP {args.hp} TWh, 수전해 {args.ely} TWh", flush=True)
    if not r.solve(n):
        print('[시험] 최적화 실패')
        return

    out = os.path.join('results', 'sector_targets_2040ess30')
    os.makedirs(out, exist_ok=True)
    row_new, reg_new = summary(n, f'B48={args.hp}, B49={args.ely}')
    row_ref, reg_ref = summary(ref, '기준(ESS-30%)')
    tab = pd.DataFrame([row_ref, row_new]).set_index('case').T
    print(tab.to_string(float_format=lambda v: f'{v:,.3f}'))
    h2 = m.build_hydrogen_balance_table(n)
    print(h2.to_string(float_format=lambda v: f'{v:,.2f}'))
    with pd.ExcelWriter(os.path.join(out, 'summary.xlsx')) as xw:
        tab.to_excel(xw, sheet_name='전국')
        h2.to_excel(xw, sheet_name='Hydrogen_Balance', index=False)
        reg_new.to_excel(xw, sheet_name='지역_신규')
        reg_ref.to_excel(xw, sheet_name='지역_기준')
    tmp = os.path.join(tempfile.gettempdir(), 'sector_targets_2040.nc')
    n.export_to_netcdf(tmp)
    import shutil
    shutil.copy2(tmp, os.path.join(out, 'sector_targets_2040.nc'))
    print(f'[시험] 저장: {out}')


if __name__ == '__main__':
    main()
