"""
반도체 수요 지역 재배치(WP2) vs 기준(REF: 재생 상한 설정, OPEN: 재생 상한 해제) 결과 비교

입력 : wp2_compare_nc/{시나리오}.nc  (한글 경로를 netCDF가 읽지 못해 영문 경로로 복사한 파일)
출력 : reports/wp2_compare_<날짜>/WP2_비교표.xlsx, fig*.png
"""
import logging
import os
from datetime import datetime

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pypsa

logging.disable(logging.WARNING)
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 9

SRC = 'wp2_compare_nc'
SCEN = [f'{c}_{y}' for y in ('2035', '2040')
        for c in ('REF', 'REFWP2', 'OPEN', 'WP2', 'WP2fix', 'OPENcap', 'WP2cap', 'OPENcapgrid', 'WP2capgrid')]
SCEN = [s for s in SCEN if os.path.exists(f'{SRC}/{s}.nc')]
CASE_LABEL = {'REF': 'REF 상한·기준배치', 'REFWP2': 'REF 상한·재배치', 'OPEN': '상한해제·기준배치',
              'WP2': '상한해제·재배치(원본)', 'WP2fix': '상한해제·재배치', 'OPENcap': '잠재량20%·기준배치',
              'WP2cap': '잠재량20%·재배치', 'OPENcapgrid': '잠재량20%+송전·기준배치', 'WP2capgrid': '잠재량20%+송전·재배치'}
LABEL = {s: f"{s.split('_')[1]} {CASE_LABEL[s.split('_')[0]]}" for s in SCEN}
REG = ['GGD', 'ICN', 'GWD', 'CBD', 'CND', 'JBD', 'GBD', 'BSN', 'GND', 'JND', 'JJD']
REGION_KR = {'BSN': '부산울산', 'CBD': '충북', 'CND': '대전세종충남', 'GBD': '대구경북', 'GGD': '서울경기',
             'GND': '경남', 'GWD': '강원', 'ICN': '인천', 'JBD': '전북', 'JJD': '제주', 'JND': '광주전남'}
HVDC_KM = {'GWD_HVDC': 230.0, 'JBD_HVDC': 250.0, 'JND_HVDC': 400.0}
OUT = os.path.join('reports', f"wp2_compare_{datetime.now().strftime('%Y%m%d')}")
os.makedirs(OUT, exist_ok=True)


def reg_of(name):
    return str(name).strip()[:3]


def analyze(s):
    n = pypsa.Network(f'{SRC}/{s}.nc')
    w = n.snapshot_weightings.generators
    g, L = n.generators, n.links
    P = n.generators_t.p.reindex(columns=g.index).fillna(0)
    E = P.mul(w, axis=0).sum() / 1e6
    loadP = n.get_switchable_as_dense('Load', 'p_set')
    lam = n.buses_t.marginal_price
    elb = [f'{r}_EL' for r in REG]

    def ltype(l):
        t = l.strip().split('_Demand_')
        return t[1] if len(t) > 1 else 'other'

    lt = pd.Series({l: ltype(l) for l in n.loads.index})
    el_loads = n.loads.index[n.loads.bus.isin(elb)]
    D = loadP.mul(w, axis=0).sum() / 1e6
    nat = {'scenario': s}
    for t in ('EL', 'DC', 'Fab', 'EV', 'H'):
        nat[f'demand_{t}_TWh'] = D[lt.index[lt == t]].sum()
    nat['demand_ELbus_TWh'] = D[el_loads].sum()
    nat['demand_all_TWh'] = D.sum()

    # 발전·설비
    for car, key in (('solar', 'PV'), ('wind', 'WT'), ('nuclear', 'NUC'), ('coal', 'COAL'), ('gas', 'GAS')):
        idx = g.index[g.carrier == car]
        nat[f'{key}_GW'] = g.loc[idx, 'p_nom_opt'].sum() / 1e3
        nat[f'{key}_new_GW'] = (g.loc[idx, 'p_nom_opt'] - g.loc[idx, 'p_nom']).sum() / 1e3
        nat[f'{key}_TWh'] = E[idx].sum()
    re = g.index[g.carrier.isin(['solar', 'wind'])]
    avail = n.generators_t.p_max_pu.reindex(columns=re).fillna(g.loc[re, 'p_max_pu']).mul(g.loc[re, 'p_nom_opt'])
    curt = (avail - P[re]).clip(lower=0).mul(w, axis=0).sum() / 1e6
    nat['curtail_TWh'] = curt.sum()
    nat['curtail_%'] = 100 * curt.sum() / (avail.mul(w, axis=0).sum().sum() / 1e6)
    slack = [x for x in g.index if x.endswith('_slack')]
    hfb = [x for x in g.index if 'H_Fallback' in x]
    nat['slack_TWh'] = E[slack].sum()
    nat['heat_fallback_TWh'] = E[hfb].sum()
    pen = (P[slack + hfb].mul(w, axis=0).sum() * g.loc[slack + hfb, 'marginal_cost']).sum()
    toEL = L.index[L.bus1.isin(elb)]
    p1 = -n.links_t.p1.reindex(columns=toEL).fillna(0)
    for car, idx in pd.Series(L.loc[toEL, 'carrier']).groupby(L.loc[toEL, 'carrier']).groups.items():
        if car != 'DC':
            nat[f'link_{car}_TWh'] = p1[idx].mul(w, axis=0).sum().sum() / 1e6
    st = n.stores[n.stores.index.str.endswith('_ESS')]
    nat['ESS_GWh'] = st.e_nom_opt.sum() / 1e3
    ch = L.index[L.index.str.endswith('_ESS_Ch')]
    nat['ESS_charge_TWh'] = n.links_t.p0[ch].clip(lower=0).mul(w, axis=0).sum().sum() / 1e6

    # 비용·가격
    nat['objective_조'] = n.objective / 1e9
    nat['penalty_조'] = pen / 1e9
    nat['cost_excl_penalty_조'] = (n.objective - pen) / 1e9
    nat['cost_per_MWh_천원'] = (n.objective - pen) / (nat['demand_all_TWh'] * 1e6)
    gc = n.global_constraints
    nat['CO2_cap_Mt'] = gc.constant.iloc[0] / 1e6
    nat['CO2_shadow_천원/t'] = -gc.mu.iloc[0]
    nat['CO2_rent_조'] = nat['CO2_shadow_천원/t'] * gc.constant.iloc[0] / 1e9
    lp = loadP[el_loads]
    lb = lam[n.loads.loc[el_loads, 'bus']].values
    pay = (lp.values * lb * w.values[:, None]).sum(axis=0)
    payS = pd.Series(pay, index=el_loads)
    nat['EL_price_loadw'] = payS.sum() / (lp.mul(w, axis=0).sum().sum())
    for t in ('EL', 'DC', 'Fab'):
        idx = [l for l in el_loads if lt[l] == t]
        q = lp[idx].mul(w, axis=0).sum().sum()
        nat[f'price_{t}_loadw'] = payS[idx].sum() / q if q > 0 else np.nan
    nat['consumer_payment_조'] = payS.sum() / 1e9
    pr = lam[elb]
    nat['price_spread_mean'] = float((pr.max(axis=1) - pr.min(axis=1)).mean())
    nat['hours_price_le1'] = float((pr <= 1).sum().mean())

    # 송전
    rows = []
    ln = n.lines
    p0 = n.lines_t.p0
    for (b0, b1), idx in ln.groupby([ln.bus0, ln.bus1]).groups.items():
        f = p0[idx].sum(axis=1)
        cap = (ln.loc[idx, 's_nom'] * ln.loc[idx, 's_max_pu']).sum()
        cong = (p0[idx].abs().div(ln.loc[idx, 's_nom'] * ln.loc[idx, 's_max_pu'])).ge(0.99).any(axis=1)
        exp = f"{reg_of(b0)}-{reg_of(b1)}_EXP"
        added = exp_cost = 0.0
        if exp in L.index:
            added = float(L.at[exp, 'p_nom_opt'])
            exp_cost = added * float(L.at[exp, 'capital_cost'])
            f = f + n.links_t.p0[exp]
            if added > 1:
                cong = cong & (n.links_t.p0[exp].abs() / added).ge(0.99)
            cap += added
        rent = (f * (lam[b1] - lam[b0]) * w).sum()
        rows.append({'scenario': s, 'corridor': f"{reg_of(b0)}-{reg_of(b1)}", 'type': 'AC', 'cap_GW': cap / 1e3,
                     'added_GW': added / 1e3, 'exp_cost_조': exp_cost / 1e9,
                     'flow_fwd_TWh': (f.clip(lower=0) * w).sum() / 1e6, 'flow_rev_TWh': (-f.clip(upper=0) * w).sum() / 1e6,
                     'flow_abs_TWh': (f.abs() * w).sum() / 1e6, 'util_%': 100 * (f.abs() * w).sum() / (cap * w.sum()),
                     'cong_h': float((cong * w).sum()), 'rent_조': rent / 1e9,
                     'TWhkm': (f.abs() * w).sum() / 1e6 * float(ln.loc[idx, 'length'].clip(lower=1).mean())})
    for l, km in HVDC_KM.items():
        if l not in L.index:
            continue
        f = n.links_t.p0[l]
        cap = L.at[l, 'p_nom_opt']
        rent = (f * (lam[L.at[l, 'bus1']] - lam[L.at[l, 'bus0']]) * w).sum()
        hv_add = cap - L.at[l, 'p_nom'] if L.at[l, 'p_nom_extendable'] else 0.0
        rows.append({'scenario': s, 'corridor': f"{reg_of(L.at[l, 'bus0'])}-{reg_of(L.at[l, 'bus1'])}", 'type': 'HVDC',
                     'cap_GW': cap / 1e3, 'added_GW': hv_add / 1e3,
                     'exp_cost_조': hv_add * float(L.at[l, 'capital_cost']) / 1e9, 'flow_fwd_TWh': (f.clip(lower=0) * w).sum() / 1e6,
                     'flow_rev_TWh': (-f.clip(upper=0) * w).sum() / 1e6, 'flow_abs_TWh': (f.abs() * w).sum() / 1e6,
                     'util_%': 100 * (f.abs() * w).sum() / (cap * w.sum()) if cap > 0 else 0,
                     'cong_h': float(((f.abs() / cap).ge(0.99) * w).sum()) if cap > 0 else 0, 'rent_조': rent / 1e9,
                     'TWhkm': (f.abs() * w).sum() / 1e6 * km})
    corr = pd.DataFrame(rows)
    nat['grid_flow_TWh'] = corr.flow_abs_TWh.sum()
    nat['grid_TWhkm'] = corr.TWhkm.sum()
    nat['grid_cong_h_sum'] = corr.cong_h.sum()
    nat['grid_rent_조'] = corr.rent_조.sum()
    nat['grid_added_GW'] = corr.added_GW.sum()
    nat['grid_exp_cost_조'] = corr.exp_cost_조.sum()

    # 지역
    net_exp = pd.Series(0.0, index=REG)
    for _, r in corr.iterrows():
        a, b = r.corridor.split('-')
        net_exp[a] += r.flow_fwd_TWh - r.flow_rev_TWh
        net_exp[b] -= r.flow_fwd_TWh - r.flow_rev_TWh
    reg = []
    for r in REG:
        rl = [l for l in el_loads if reg_of(l) == r]
        q = lp[rl].mul(w, axis=0).sum().sum() / 1e6
        gi = g.index[(g.bus == f'{r}_EL')]
        gr = [x for x in gi if g.at[x, 'carrier'] in ('solar', 'wind')]
        reg.append({'scenario': s, 'region': r, 'load_TWh': q,
                    'Fab_TWh': D[[l for l in rl if lt[l] == 'Fab']].sum(),
                    'PV_GW': g.loc[[x for x in gr if g.at[x, 'carrier'] == 'solar'], 'p_nom_opt'].sum() / 1e3,
                    'WT_GW': g.loc[[x for x in gr if g.at[x, 'carrier'] == 'wind'], 'p_nom_opt'].sum() / 1e3,
                    'RE_TWh': E[gr].sum(), 'curtail_TWh': curt.reindex(gr).sum(),
                    'gen_TWh': E[[x for x in gi if x not in slack]].sum(),
                    'slack_TWh': E.reindex([f'{r}_slack']).fillna(0).sum(),
                    'net_export_TWh': net_exp[r], 'import_dep_%': 100 * max(-net_exp[r], 0) / q if q else 0,
                    'RE_selfsuff_%': 100 * E[gr].sum() / q if q else 0,
                    'price_loadw': payS[rl].sum() / (lp[rl].mul(w, axis=0).sum().sum()) if q else np.nan})
    return nat, corr, pd.DataFrame(reg)


nats, corrs, regs = [], [], []
for s in SCEN:
    a, b, c = analyze(s)
    nats.append(a); corrs.append(b); regs.append(c)
    print('done', s, flush=True)
nat = pd.DataFrame(nats).set_index('scenario').T
corr = pd.concat(corrs)
reg = pd.concat(regs)

with pd.ExcelWriter(os.path.join(OUT, 'WP2_비교표.xlsx')) as wr:
    nat.rename(columns=LABEL).to_excel(wr, sheet_name='전국지표')
    for col in ('load_TWh', 'Fab_TWh', 'PV_GW', 'WT_GW', 'RE_TWh', 'curtail_TWh', 'slack_TWh', 'net_export_TWh',
                'import_dep_%', 'RE_selfsuff_%', 'price_loadw'):
        reg.pivot_table(index='region', columns='scenario', values=col)[SCEN].rename(
            columns=LABEL, index=REGION_KR).to_excel(wr, sheet_name=f'지역_{col}'[:31])
    for col in ('flow_abs_TWh', 'util_%', 'cong_h', 'rent_조'):
        corr.pivot_table(index=['type', 'corridor'], columns='scenario', values=col)[SCEN].rename(
            columns=LABEL).to_excel(wr, sheet_name=f'송전_{col}'[:31])
nat.to_csv(os.path.join(OUT, 'national.csv'), encoding='utf-8-sig')
corr.to_csv(os.path.join(OUT, 'corridor.csv'), index=False, encoding='utf-8-sig')
reg.to_csv(os.path.join(OUT, 'regional.csv'), index=False, encoding='utf-8-sig')
pd.set_option('display.width', 300)
pd.set_option('display.max_rows', 200)
print(nat.round(2).to_string())
