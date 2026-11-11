"""
2035년 민감도 분석 보고서(Word) + 그림 + 표 생성

선행: python analyze_2035.py <민감도 결과폴더> <MGA 결과폴더 ...> --out reports/sens2035_<날짜>
사용법:
  python make_2035_report.py [--dir reports/sens2035_...] [--ref2040 reports/mga_...] [--ref2040-cap reports/mga_cap20_...]
"""
import argparse
import glob
import os
from datetime import datetime

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from run_regional_cost_sweep import TECH_POTENTIAL_GW

plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 9

ap = argparse.ArgumentParser()
ap.add_argument('--dir', default=sorted(glob.glob('reports/sens2035_*'))[-1])
ap.add_argument('--ref2040', default='reports/mga_20261103')
ap.add_argument('--ref2040-cap', default='reports/mga_cap20_20261103')
args = ap.parse_args()
OUT = args.dir

nat = pd.read_csv(f'{OUT}/national.csv', index_col=0)
reg = pd.read_csv(f'{OUT}/regional.csv')
sw = pd.read_csv(f'{OUT}/sweep_summary.csv').set_index('case')
r40 = pd.read_csv(f'{args.ref2040}/mga_national.csv', index_col=0)
r40c = pd.read_csv(f'{args.ref2040_cap}/mga_national.csv', index_col=0)
reg40 = pd.read_csv(f'{args.ref2040}/mga_regional.csv')

REGION_KR = {'BSN': '부산울산', 'CBD': '충북', 'CND': '대전세종충남', 'GBD': '대구경북', 'GGD': '서울경기',
             'GND': '경남', 'GWD': '강원', 'ICN': '인천', 'JBD': '전북', 'JJD': '제주', 'JND': '광주전남'}
REG = ['GGD', 'ICN', 'GWD', 'CBD', 'CND', 'JBD', 'GBD', 'BSN', 'GND', 'JND', 'JJD']
CASE_KR = {'base': '기준', 'cap20': '잠재량 20% 상한', 'EL_m05': '일반 −5%', 'EL_m10': '일반 −10%',
           'DC_m25': 'DC −25%', 'Fab_m25': 'Fab −25%', 'DCFab_m25': 'DC+Fab −25%', 'DCFab_m50': 'DC+Fab −50%',
           'B0': 'B0 (MGA 기준)', 'MGA-1': 'MGA 1%', 'MGA-2': 'MGA 2%'}
DEM = ['EL_m05', 'EL_m10', 'DC_m25', 'Fab_m25', 'DCFab_m25', 'DCFab_m50']
MGA = [c for c in ('MGA-1', 'MGA-2') if c in nat.index]
NINE = ['cap20'] + DEM + MGA
C_PV, C_WT, C_B, C_PEN, C_GAS, C_NU, C_COAL, C_CHP = ('#F2B134', '#3C8DBC', '#1F4E79', '#C0392B', '#E07B39',
                                                      '#6C757D', '#3E2723', '#8E5EA2')
C_EL, C_DC, C_FAB = '#5B7DB1', '#E07B39', '#8E5EA2'

LIM = {t: {r: max(0.2 * TECH_POTENTIAL_GW[t][r], float(reg[(reg.case == 'base') & (reg.region == r)][f'{t}_exist_GW'].iloc[0]))
           for r in REG} for t in ('PV', 'WT')}
R = {c: reg[reg.case == c].set_index('region').reindex(REG) for c in nat.index}
R40 = {c: reg40[reg40.case == c].set_index('region').reindex(REG) for c in reg40.case.unique()}


def V(c, k):
    return float(nat.at[c, k])


def L(c):
    return CASE_KR.get(c, c)


def Lp(c):
    return L(c).replace('−', '-')


def d(c, k, base='base'):
    return V(c, k) - V(base, k)


C0 = V('base', 'cost_sys')
B0 = V('B0', 'cost_sys')
RE0 = V('base', 'RE_GW')
for c in DEM:
    nat.at[c, 'd_demand'] = V('base', 'demand_power_TWh') - V(c, 'demand_power_TWh')
    nat.at[c, 'd_cost'] = C0 - V(c, 'cost_sys')
    nat.at[c, 'avoided'] = nat.at[c, 'd_cost'] * 1e9 / (nat.at[c, 'd_demand'] * 1e6)
AV = nat.loc[DEM, 'avoided']


def mga_inc(c):
    return 100 * (V(c, 'cost_sys') / B0 - 1)


def mga_dre(c):
    return V(c, 'RE_GW') - V('B0', 'RE_GW')


def pot_pct(c, t, r):
    return R[c].at[r, f'{t}_GW'] / TECH_POTENTIAL_GW[t][r] * 100


figs = {}


def save(fig, name, rect=(0, 0, 1, 1)):
    p = os.path.join(OUT, name)
    fig.tight_layout(rect=rect)
    fig.savefig(p, dpi=200)
    plt.close(fig)
    figs[name] = p


# ───────────── 그림 ─────────────
# fig0 기준 시나리오
fig, ax = plt.subplots(1, 3, figsize=(13, 3.7), gridspec_kw={'width_ratios': [1, 1, 1.6]})
sup = [('원자력', 'nuclear_TWh', C_NU), ('석탄', 'coal_TWh', C_COAL), ('LNG 발전', 'LNG_power_TWh', C_GAS),
       ('CHP', 'CHP_power_TWh', C_CHP), ('연료전지', 'fuelcell_power_TWh', '#2E7D32'),
       ('PV', 'PV_TWh', C_PV), ('풍력', 'WT_TWh', C_WT)]
bot = 0
for nm, k, col in sup:
    v = V('base', k)
    ax[0].bar(0, v, bottom=bot, color=col, width=0.5, label=f'{nm} {v:.0f}')
    bot += v
ax[0].set_xlim(-0.6, 1.6); ax[0].set_xticks([]); ax[0].set_ylabel('TWh')
ax[0].legend(fontsize=7.5, loc='center right'); ax[0].set_title('(a) 기준 발전량 구성')
dem = [('일반', 'demand_EL_TWh', C_EL), ('데이터센터', 'demand_DC_TWh', C_DC), ('팹', 'demand_Fab_TWh', C_FAB),
       ('히트펌프', 'HP_elec_TWh', '#9E9E9E'), ('수전해', 'electrolyser_elec_TWh', '#80CBC4'), ('EV 충전', 'EV_charge_TWh', '#BCAAA4')]
bot = 0
for nm, k, col in dem:
    v = V('base', k)
    ax[1].bar(0, v, bottom=bot, color=col, width=0.5, label=f'{nm} {v:.0f}')
    bot += v
ax[1].set_xlim(-0.6, 1.6); ax[1].set_xticks([]); ax[1].legend(fontsize=7.5, loc='center right')
ax[1].set_title('(b) 기준 전력 소비 구성 (TWh)')
xx = np.arange(len(REG))
ax[2].bar(xx - 0.2, R['base'].PV_GW, 0.4, color=C_PV, label='PV')
ax[2].bar(xx + 0.2, R['base'].WT_GW, 0.4, color=C_WT, label='풍력')
ax[2].bar(xx - 0.2, R['base'].PV_exist_GW, 0.4, color='none', edgecolor='k', hatch='///', lw=0.5, label='기존 설비')
ax[2].bar(xx + 0.2, R['base'].WT_exist_GW, 0.4, color='none', edgecolor='k', hatch='///', lw=0.5)
ax[2].scatter(xx - 0.2, [TECH_POTENTIAL_GW['PV'][r] * 0.2 for r in REG], marker='_', s=120, color=C_PEN, zorder=4,
              label='잠재량 20%')
ax[2].scatter(xx + 0.2, [TECH_POTENTIAL_GW['WT'][r] * 0.2 for r in REG], marker='_', s=120, color=C_PEN, zorder=4)
ax[2].set_xticks(xx, [REGION_KR[r] for r in REG], rotation=40, ha='right', fontsize=7.5)
ax[2].set_ylabel('GW'); ax[2].legend(fontsize=7.5); ax[2].grid(axis='y', alpha=0.3)
ax[2].set_title('(c) 기준 지역별 PV·풍력 설비')
save(fig, 'fig0_base.png')

# fig1 전체 요약
cases = ['base'] + NINE
fig, ax = plt.subplots(1, 2, figsize=(13, 3.9))
xx = np.arange(len(cases))
pv = [V(c, 'PV_GW') for c in cases]; wt = [V(c, 'WT_GW') for c in cases]
ax[0].bar(xx, pv, color=C_PV, label='PV'); ax[0].bar(xx, wt, bottom=pv, color=C_WT, label='풍력')
for i in range(len(cases)):
    ax[0].text(i, pv[i] + wt[i] + 4, f'{pv[i] + wt[i]:.0f}', ha='center', fontsize=7.5)
ax[0].set_xticks(xx, [Lp(c) for c in cases], rotation=35, ha='right', fontsize=7.5)
ax[0].set_ylabel('GW'); ax[0].set_ylim(0, 340); ax[0].legend(fontsize=8, loc='lower left'); ax[0].grid(axis='y', alpha=0.3)
ax[0].set_title('(a) PV·풍력 설비')
dc = [V(c, 'cost_sys') - C0 for c in NINE]
cols = [C_PEN if v > 0 else C_B for v in dc]
ax[1].bar(np.arange(len(NINE)), dc, color=cols)
for i, v in enumerate(dc):
    ax[1].text(i, v + (0.3 if v >= 0 else -0.9), f'{v:+.2f}', ha='center', fontsize=7.5)
ax[1].axhline(0, color='k', lw=0.6)
ax[1].set_xticks(np.arange(len(NINE)), [Lp(c) for c in NINE], rotation=35, ha='right', fontsize=7.5)
ax[1].set_ylabel('기준 대비 시스템 비용 변화 (조원/년)'); ax[1].grid(axis='y', alpha=0.3)
ax[1].set_title('(b) 시스템 비용 변화')
save(fig, 'fig1_overview.png')

# fig2 잠재량 상한
fig, ax = plt.subplots(1, 2, figsize=(12, 4.0), sharey=True)
yy = np.arange(len(REG))[::-1]
for k, t in enumerate(('PV', 'WT')):
    ax[k].barh(yy + 0.2, R['base'][f'{t}_GW'], 0.4, color='#BDD7EE', label='기준')
    ax[k].barh(yy - 0.2, R['cap20'][f'{t}_GW'], 0.4, color=C_B, label='잠재량 20% 상한')
    ax[k].scatter([LIM[t][r] for r in REG], yy, marker='|', s=260, color=C_PEN, zorder=4, label='상한')
    ax[k].scatter([TECH_POTENTIAL_GW[t][r] for r in REG], yy, marker='x', s=20, color='k', zorder=4, label='기술적 잠재량 100%')
    ax[k].set_title(f"({'ab'[k]}) {'태양광' if t == 'PV' else '풍력'} 지역별 설비 (GW)"); ax[k].grid(axis='x', alpha=0.3)
ax[0].set_xlim(0, 80); ax[1].set_xlim(0, 50)
ax[0].set_yticks(yy, [REGION_KR[r] for r in REG])
h, l_ = ax[0].get_legend_handles_labels()
fig.legend(h, l_, fontsize=8, loc='lower center', ncol=4, bbox_to_anchor=(0.5, -0.01))
save(fig, 'fig2_cap20.png', rect=(0, 0.07, 1, 1))

# fig3 수요 감축
colmap = {'EL_m05': C_EL, 'EL_m10': C_EL, 'DC_m25': C_DC, 'Fab_m25': C_FAB, 'DCFab_m25': '#555555', 'DCFab_m50': '#555555'}
fig, ax = plt.subplots(1, 2, figsize=(11.5, 3.8))
for c in DEM:
    ax[0].scatter(nat.at[c, 'd_demand'], nat.at[c, 'd_cost'], color=colmap[c], s=40, zorder=3)
    ax[0].annotate(Lp(c), (nat.at[c, 'd_demand'], nat.at[c, 'd_cost']), textcoords='offset points',
                   xytext={'DCFab_m25': (4, -11), 'EL_m10': (-50, 4)}.get(c, (4, 3)), fontsize=7.5)
xm = np.linspace(0, nat.loc[DEM, 'd_demand'].max() * 1.05, 10)
ax[0].plot(xm, xm * AV.mean() / 1e3, '--', color='gray', lw=1, label=f'평균 {AV.mean():.0f} 천원/MWh')
ax[0].set_xlabel('전력수요 감소 (TWh)'); ax[0].set_ylabel('시스템 비용 감소 (조원/년)'); ax[0].grid(alpha=0.3)
ax[0].legend(fontsize=8, loc='upper left'); ax[0].set_title('(a) 수요 감소량과 비용 감소')
xx = np.arange(len(DEM))
ax[1].bar(xx - 0.2, [d(c, 'PV_GW') for c in DEM], 0.4, color=C_PV, label='PV')
ax[1].bar(xx + 0.2, [d(c, 'WT_GW') for c in DEM], 0.4, color=C_WT, label='풍력')
ax[1].axhline(0, color='k', lw=0.6)
ax[1].set_xticks(xx, [Lp(c) for c in DEM], rotation=25, ha='right', fontsize=8)
ax[1].set_ylabel('기준 대비 설비 변화 (GW)'); ax[1].legend(fontsize=8, loc='lower left'); ax[1].grid(axis='y', alpha=0.3)
ax[1].set_title('(b) 수요 감축 케이스별 PV·풍력 변화')
save(fig, 'fig3_demand.png')

# fig4 지역별 변화 히트맵 (수요 감축 + 상한 + MGA)
HM = DEM + MGA
keys = [(t, r) for t in ('PV', 'WT') for r in REG]
mat = np.array([[R[c].at[r, f'{t}_GW'] - R['base' if not c.startswith('MGA') else 'B0'].at[r, f'{t}_GW'] for c in HM]
                for t, r in keys])
keep = np.abs(mat).max(axis=1) > 0.5
mat = mat[keep]; kk = [k for k, f in zip(keys, keep) if f]
fig, ax = plt.subplots(figsize=(9.5, 0.36 * len(kk) + 1.3))
vm = np.abs(mat).max()
im = ax.imshow(mat, cmap='RdBu', vmin=-vm, vmax=vm, aspect='auto')
ax.set_yticks(range(len(kk)), [f"{REGION_KR[r]} {'PV' if t == 'PV' else '풍력'} (기준 {R['base'].at[r, f'{t}_GW']:.1f})"
                               for t, r in kk], fontsize=8)
ax.set_xticks(range(len(HM)), [Lp(c) for c in HM], rotation=30, ha='right', fontsize=8)
for i in range(mat.shape[0]):
    for j in range(mat.shape[1]):
        ax.text(j, i, f'{mat[i, j]:+.1f}', ha='center', va='center', fontsize=7,
                color='white' if abs(mat[i, j]) > vm * 0.6 else 'black')
fig.colorbar(im, ax=ax, shrink=0.9, label='기준(MGA는 B0) 대비 변화 (GW)')
ax.set_title('지역별 PV·풍력 설비 변화 (변화 0.5 GW 이상 항목)')
save(fig, 'fig4_regional_change.png')

# fig5 MGA
fig, ax = plt.subplots(1, 2, figsize=(12.5, 3.9), gridspec_kw={'width_ratios': [1, 1.6]})
x35 = [0.0] + [mga_inc(c) for c in MGA]; y35 = [0.0] + [100 * mga_dre(c) / V('B0', 'RE_GW') for c in MGA]
m40 = [c for c in ('MGA-0.5', 'MGA-1', 'MGA-2', 'MGA-3') if c in r40.index]
b40c = float(r40.at['B0', 'cost_sys']); b40r = float(r40.at['B0', 'RE_GW'])
x40 = [0.0] + [100 * (float(r40.at[c, 'cost_sys']) / b40c - 1) for c in m40]
y40 = [0.0] + [100 * (float(r40.at[c, 'RE_GW']) / b40r - 1) for c in m40]
ax[0].plot(x40, y40, '--o', color='#A6A6A6', ms=4, label='2040년 (상한 해제)')
ax[0].plot(x35, y35, '-o', color=C_B, label='2035년')
for c, x_, y_ in zip(MGA, x35[1:], y35[1:]):
    ax[0].annotate(f'{mga_dre(c):+.1f} GW', (x_, y_), textcoords='offset points', xytext=(5, 4), fontsize=7.5, color=C_B)
ax[0].set_xlabel('비용 증가율 (%, 각 연도 B0 대비)'); ax[0].set_ylabel('PV+풍력 총설비 변화 (%)')
ax[0].grid(alpha=0.3); ax[0].legend(fontsize=8); ax[0].set_title('(a) 비용 증가 허용률별 최소 재생설비')
items = [('PV 발전', 'PV_TWh'), ('풍력 발전', 'WT_TWh'), ('출력제한', 'curtail_TWh'), ('LNG 발전', 'LNG_power_TWh'),
         ('CHP 발전', 'CHP_power_TWh'), ('석탄', 'coal_TWh'), ('원자력', 'nuclear_TWh'),
         ('수전해 전력', 'electrolyser_elec_TWh'), ('연료전지', 'fuelcell_power_TWh'), ('EV 충전', 'EV_charge_TWh'),
         ('ESS 충전', 'ESS_charge_TWh')]
xx = np.arange(len(items)); wd = 0.8 / max(len(MGA), 1)
for j, (c, col) in enumerate(zip(MGA, ('#2E75B6', C_B))):
    ax[1].bar(xx - 0.4 + wd * (j + 0.5), [d(c, k, 'B0') for _, k in items], wd, color=col, label=Lp(c))
ax[1].axhline(0, color='k', lw=0.6)
ax[1].set_xticks(xx, [n_ for n_, _ in items], rotation=30, ha='right', fontsize=8)
ax[1].set_ylabel('B0 대비 변화 (TWh)'); ax[1].legend(fontsize=8); ax[1].grid(axis='y', alpha=0.3)
ax[1].set_title('(b) MGA 대안의 연간 전력 수급 변화')
save(fig, 'fig5_mga.png')


# ───────────── 표 ─────────────
def summary_table():
    rows = []
    for c in ['base'] + NINE:
        ref = 'B0' if c.startswith('MGA') else 'base'
        rows.append({'케이스': L(c), '전력수요 (TWh)': V(c, 'demand_power_TWh'), 'PV (GW)': V(c, 'PV_GW'),
                     '풍력 (GW)': V(c, 'WT_GW'), 'PV+풍력 (GW)': V(c, 'RE_GW'), '출력제한률 (%)': V(c, 'curtail_%'),
                     '시스템 비용 (조원)': V(c, 'cost_sys'),
                     '비용 변화 (%)': 100 * (V(c, 'cost_sys') / V(ref, 'cost_sys') - 1),
                     'CO2 잠재가격 (천원/t)': V(c, 'CO2_price'), 'HVDC (TWh)': V(c, 'HVDC_TWh')})
    return pd.DataFrame(rows)


def cap_table():
    items = [('시스템 비용 (조원)', 'cost_sys'), ('PV (GW)', 'PV_GW'), ('풍력 (GW)', 'WT_GW'), ('PV+풍력 (GW)', 'RE_GW'),
             ('출력제한률 (%)', 'curtail_%'), ('PV 이용률 (%)', 'PV_CF_%'), ('풍력 이용률 (%)', 'WT_CF_%'),
             ('PV 투자비 (조원)', 'c_capex_PV'), ('풍력 투자비 (조원)', 'c_capex_WT'), ('가스 연료비 (조원)', 'c_opex_gas'),
             ('LNG 발전 (TWh)', 'LNG_power_TWh'), ('HVDC 수송량 (TWh)', 'HVDC_TWh'), ('송전 혼잡 (선로·시간)', 'line_congestion_h'),
             ('CO2 잠재가격 (천원/t)', 'CO2_price'), ('부하가중 전력가격 (천원/MWh)', 'EL_price_loadw')]
    return pd.DataFrame([{'항목': a, '기준': V('base', k), '잠재량 20% 상한': V('cap20', k), '차이': d('cap20', k)}
                         for a, k in items])


def demand_table():
    rows = []
    for c in ['base'] + DEM:
        rows.append({'케이스': L(c), '일반 (TWh)': V(c, 'demand_EL_TWh'), 'DC (TWh)': V(c, 'demand_DC_TWh'),
                     'Fab (TWh)': V(c, 'demand_Fab_TWh'),
                     '수요 감소 (TWh)': nat.at[c, 'd_demand'] if c != 'base' else 0.0,
                     'PV (GW)': V(c, 'PV_GW'), '풍력 (GW)': V(c, 'WT_GW'),
                     '재생 발전 (TWh)': V(c, 'PV_TWh') + V(c, 'WT_TWh'), '출력제한률 (%)': V(c, 'curtail_%'),
                     'LNG 발전 (TWh)': V(c, 'LNG_power_TWh'), '석탄 (TWh)': V(c, 'coal_TWh'),
                     '시스템 비용 (조원)': V(c, 'cost_sys'),
                     '비용 감소 (조원)': nat.at[c, 'd_cost'] if c != 'base' else 0.0,
                     'MWh당 회피비용 (천원)': nat.at[c, 'avoided'] if c != 'base' else np.nan,
                     'CO2 잠재가격 (천원/t)': V(c, 'CO2_price')})
    return pd.DataFrame(rows)


def mga_table():
    rows = []
    for c in ['B0'] + MGA:
        rows.append({'시나리오': L(c), '비용 증가 허용률': '—' if c == 'B0' else f"{float(nat.at[c, 'eps_pct']):g}%",
                     '시스템 비용 (조원)': V(c, 'cost_sys'), '실제 비용 증가율 (%)': 0.0 if c == 'B0' else mga_inc(c),
                     '태양광 GW': V(c, 'PV_GW'), '풍력 GW': V(c, 'WT_GW'),
                     'B0 대비 재생 감소 GW': -mga_dre(c),
                     'ESS GW/GWh': f"{V(c, 'ESS_GW'):.1f} / {V(c, 'ESS_GWh'):.0f}", '출력제한율 (%)': V(c, 'curtail_%')})
    return pd.DataFrame(rows)


def balance_table():
    items = [('PV 발전', 'PV_TWh'), ('풍력 발전', 'WT_TWh'), ('출력제한', 'curtail_TWh'), ('원자력', 'nuclear_TWh'),
             ('석탄', 'coal_TWh'), ('LNG 발전', 'LNG_power_TWh'), ('CHP 발전', 'CHP_power_TWh'), ('CHP 열', 'CHP_heat_TWh'),
             ('가스 연료 총량', 'gas_fuel_TWh'), ('연료전지 발전', 'fuelcell_power_TWh'), ('히트펌프 전력', 'HP_elec_TWh'),
             ('수전해 전력', 'electrolyser_elec_TWh'), ('EV 충전', 'EV_charge_TWh'), ('V2G 방전', 'V2G_TWh'),
             ('ESS 충전', 'ESS_charge_TWh'), ('전력 소비 합계', 'elec_demand_total_TWh'), ('HVDC 수송량', 'HVDC_TWh'),
             ('CO2 배출 (Mt)', 'CO2_Mt'), ('CO2 잠재가격 (천원/t)', 'CO2_price'), ('부하가중 전력가격 (천원/MWh)', 'EL_price_loadw')]
    df = pd.DataFrame({'항목 (TWh)': [a for a, _ in items]})
    for c in ['base'] + NINE:
        df[L(c)] = [V(c, k) for _, k in items]
    return df


def regional_table():
    rows = []
    for r in REG:
        row = {'지역': REGION_KR[r]}
        for t, tk in (('PV', 'PV'), ('WT', '풍력')):
            row[f'{tk} 잠재량'] = TECH_POTENTIAL_GW[t][r]
            row[f'{tk} 기존'] = R['base'].at[r, f'{t}_exist_GW']
            for c in ['base', 'cap20', 'EL_m10', 'DCFab_m50'] + MGA[-1:]:
                row[f'{tk} {L(c)}'] = R[c].at[r, f'{t}_GW']
        rows.append(row)
    return pd.DataFrame(rows)


def y2040_table():
    rows = [['기준 PV / 풍력 (GW)', f"{V('base', 'PV_GW'):.1f} / {V('base', 'WT_GW'):.1f}",
             f"{float(r40.at['B0', 'PV_GW']):.1f} / {float(r40.at['B0', 'WT_GW']):.1f}"],
            ['기준 출력제한률 (%)', f"{V('base', 'curtail_%'):.1f}", f"{float(r40.at['B0', 'curtail_%']):.1f}"],
            ['기준 시스템 비용 (조원)', f'{C0:.2f}', f"{float(r40.at['B0', 'cost_sys']):.2f}"],
            ['CO2 상한 (Mt) / 잠재가격 (천원/t)', f"{V('base', 'CO2_cap_Mt'):.1f} / {V('base', 'CO2_price'):.0f}",
             f"{float(r40.at['B0', 'CO2_Mt']):.1f} / {float(r40.at['B0', 'CO2_price']):.0f}"],
            ['잠재량 20% 상한: 비용 / 재생설비 변화',
             f"+{d('cap20', 'cost_sys'):.2f}조 (+{100 * d('cap20', 'cost_sys') / C0:.2f}%) / {d('cap20', 'RE_GW'):+.1f} GW",
             f"+{float(r40c.at['B0', 'cost_sys'] - r40.at['B0', 'cost_sys']):.2f}조 "
             f"(+{100 * float(r40c.at['B0', 'cost_sys'] / r40.at['B0', 'cost_sys'] - 1):.2f}%) / "
             f"{float(r40c.at['B0', 'RE_GW'] - r40.at['B0', 'RE_GW']):+.1f} GW"]]
    for c in MGA:
        if c in r40.index:
            rows.append([f'MGA {c[4:]}%: 재생 감소 (GW, B0 대비 %)',
                         f"{-mga_dre(c):.1f} ({-100 * mga_dre(c) / V('B0', 'RE_GW'):.1f}%)",
                         f"{b40r - float(r40.at[c, 'RE_GW']):.1f} ({100 * (1 - float(r40.at[c, 'RE_GW']) / b40r):.1f}%)"])
    return pd.DataFrame(rows, columns=['항목', '2035년', '2040년'])


TBL = {'표1_전체요약': summary_table(), '표2_잠재량상한': cap_table(), '표3_수요감축': demand_table(),
       '표4_MGA': mga_table(), '표5_수급': balance_table(), '표6_지역별': regional_table(), '표7_2040비교': y2040_table()}
with pd.ExcelWriter(os.path.join(OUT, '2035_민감도분석_표.xlsx')) as xw:
    for k, v in TBL.items():
        v.to_excel(xw, sheet_name=k, index=False)
    nat.T.to_excel(xw, sheet_name='전국_전체지표')
    reg.to_excel(xw, sheet_name='지역_전체지표', index=False)

# ───────────── Word ─────────────
doc = Document()
st = doc.styles['Normal']
st.font.name = 'Malgun Gothic'; st.font.size = Pt(10)
st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
for sname in ('Heading 1', 'Heading 2', 'Heading 3', 'Title'):
    hs = doc.styles[sname]
    hs.font.name = 'Malgun Gothic'
    hs.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
    hs.font.color.rgb = RGBColor(0x1F, 0x3A, 0x5F)
for sec in doc.sections:
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin = sec.bottom_margin = Cm(2.0)


def P(text, bold=False, size=None, italic=False, align=None):
    p = doc.add_paragraph()
    r = p.add_run(text); r.bold = bold; r.italic = italic
    if size:
        r.font.size = Pt(size)
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(3)
    return p


def B(items):
    for it in items:
        p = doc.add_paragraph(style='List Bullet')
        if isinstance(it, tuple):
            r = p.add_run(it[0]); r.bold = True
            p.add_run(it[1])
        else:
            p.add_run(it)
        p.paragraph_format.space_after = Pt(1)


def T(df, font=8, caption=None, fmt='{:,.1f}', fmts=None):
    if caption:
        P(caption, bold=True, size=9)
    t = doc.add_table(rows=1, cols=len(df.columns))
    t.style = 'Light Grid Accent 1'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, c in enumerate(df.columns):
        cell = t.rows[0].cells[j]; cell.text = ''
        r = cell.paragraphs[0].add_run(str(c)); r.bold = True; r.font.size = Pt(font)
    for _, row in df.iterrows():
        cells = t.add_row().cells
        for j, c in enumerate(df.columns):
            v = row[c]
            if isinstance(v, (float, np.floating)):
                txt = '–' if pd.isna(v) else (fmts or {}).get(c, fmt).format(v)
            else:
                txt = str(v)
            cells[j].text = ''
            rr = cells[j].paragraphs[0].add_run(txt); rr.font.size = Pt(font)
            cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if j else WD_ALIGN_PARAGRAPH.LEFT
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def F(name, caption, width=16.5):
    doc.add_picture(figs[name], width=Cm(width))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    P(caption, italic=True, size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)


def H(text, lv):
    doc.add_heading(text, lv)


FM = {'시스템 비용 (조원)': '{:,.2f}', '비용 변화 (%)': '{:+.2f}', '비용 감소 (조원)': '{:,.2f}',
      'MWh당 회피비용 (천원)': '{:,.0f}', 'CO2 잠재가격 (천원/t)': '{:,.0f}', '실제 비용 증가율 (%)': '{:,.2f}',
      '차이': '{:+,.2f}', '기준': '{:,.2f}', '잠재량 20% 상한': '{:,.2f}'}
mL = MGA[-1] if MGA else None
red40_txt = ', '.join(f"{b40r - float(r40.at[c, 'RE_GW']):.1f}" for c in MGA if c in r40.index)
ggd_pot = TECH_POTENTIAL_GW['WT']['GGD']
icn_pot = TECH_POTENTIAL_GW['PV']['ICN']
moved_out = [r for r in REG for t in ('PV', 'WT') if R['base'].at[r, f'{t}_GW'] > LIM[t][r] + 0.5]
cap_gain = sorted([(R['cap20'].at[r, f'{t}_GW'] - R['base'].at[r, f'{t}_GW'], r, t) for r in REG for t in ('PV', 'WT')],
                  reverse=True)
cap_gain_txt = ', '.join(f"{REGION_KR[r]} {'PV' if t == 'PV' else '풍력'} {v:+.1f}" for v, r, t in cap_gain[:5])
cap_loss_txt = ', '.join(f"{REGION_KR[r]} {'PV' if t == 'PV' else '풍력'} {v:+.1f}" for v, r, t in sorted(cap_gain)[:4])

doc.add_heading('PyPSA 전력계통 모형 2035년 민감도 분석 결과', 0)
P('2035년 재생에너지 보급 상한 해제 시나리오: 잠재량 상한, 전력수요 감축, 비용 제한 최소 재생설비 대안(MGA)', size=12)
P(f'작성일: {datetime.now():%Y-%m-%d}   |   모형: PyPSA 0.31.1 + linopy + CPLEX Barrier, 8,784시간 단일연도', size=9)

H('개요', 1)
P('2040년에 수행한 민감도 분석 중 일부를 2035년 조건으로 다시 수행하였다. 기준은 현재 interface에 입력된 2035년 수치이며, '
  '재생에너지 보급 상한은 해제되어 있다(지역별 PV·풍력 추가 보급 제한 없음). 입력 파일과 PyPSA_GUI.py는 수정하지 않았고, '
  '각 케이스의 조건 변경은 네트워크 생성 후 메모리에서만 적용하였다. 방법은 2040년 분석과 같다.')
T(pd.DataFrame([
    ['1', '잠재량 상한', '지역별 PV·풍력 상한 = max(기술적 잠재량 × 20%, 기존 설비)'],
    ['2~3', '일반 수요', '{지역}_Demand_EL −5%, −10% (시간대별 값 전체에 같은 비율)'],
    ['4~7', 'DC·Fab 수요', 'DC −25%, Fab −25%, DC+Fab 동시 −25%, −50%'],
    ['8~9', 'MGA', '비용 증가 허용률 1%, 2% (기준 비용최소해 B0 대비). 1단계 PV·풍력 신규 최소화 → 2단계 비용 최소화'],
], columns=['No.', '구분', '조건']), font=8, caption='표 0-1. 분석 케이스 (총 9개)')
T(pd.DataFrame([
    ['기준 입력', 'interface.xlsx 현재 입력 (2035년). 석탄 26.9 GW 존속, CO2 상한 '
     f"{V('base', 'CO2_cap_Mt'):.1f} Mt"],
    ['기준 전력수요', f"일반 {V('base', 'demand_EL_TWh'):.1f} + DC {V('base', 'demand_DC_TWh'):.1f} + Fab "
     f"{V('base', 'demand_Fab_TWh'):.1f} = {V('base', 'demand_power_TWh'):.1f} TWh (히트펌프·수전해·EV 충전 제외)"],
    ['비용 정의', 'PyPSA 목적함수(신규 설비 연간화 투자비 + 전체 운전비, 기존 설비 고정비 제외) − 열보조공급 페널티. '
     'MGA와 같은 정의를 모든 케이스에 적용'],
    ['MGA 기준(B0)', f"기준과 같은 입력에서 전력 slack을 0으로 고정해 다시 계산. 비용 {B0:.3f}조원으로 기준({C0:.3f}조원)과 "
     f"{100 * (B0 / C0 - 1):+.2f}% 차이(기준의 slack 발전량도 0)"],
], columns=['항목', '내용']), font=8, caption='표 0-2. 공통 설정')

H('핵심 결과', 2)
B([('1. 2035년 기준: ', f"PV {V('base', 'PV_GW'):.1f} GW, 풍력 {V('base', 'WT_GW'):.1f} GW(합계 {RE0:.1f} GW), 출력제한률 "
    f"{V('base', 'curtail_%'):.1f}%, 시스템 비용 {C0:.2f}조원/년. CO2 상한이 구속되고(잠재가격 {V('base', 'CO2_price'):.0f}천원/t), "
    f"석탄은 설비가 남아 있지만 발전량은 {V('base', 'coal_TWh'):.1f} TWh에 그친다. 인천 PV({R['base'].at['ICN', 'PV_GW']:.1f} GW)는 "
    f"기술적 잠재량({icn_pot:.0f} GW)을 넘고, 서울경기 풍력({R['base'].at['GGD', 'WT_GW']:.1f} GW)은 잠재량의 "
    f"{pot_pct('base', 'WT', 'GGD'):.0f}%이다."),
   ('2. 잠재량 20% 상한: ', f"비용 +{d('cap20', 'cost_sys'):.2f}조원(+{100 * d('cap20', 'cost_sys') / C0:.2f}%), 재생설비 "
    f"{d('cap20', 'RE_GW'):+.1f} GW로 총량은 거의 같다. 대신 입지가 크게 바뀐다({cap_loss_txt} GW → {cap_gain_txt} GW). "
    f"원거리 설비가 늘어 HVDC 수송량이 {V('base', 'HVDC_TWh'):.0f} → {V('cap20', 'HVDC_TWh'):.0f} TWh로 는다."),
   ('3. 수요 감축: ', f"수요 1 MWh를 줄이면 시스템 비용이 {AV.min():.0f}~{AV.max():.0f}천원 줄어든다(평균 {AV.mean():.0f}천원/MWh). "
    '감축 대상(일반·DC·Fab)과 규모에 거의 무관하다. 감축분은 주로 서울경기·충북 풍력 신규 설치 회피로 나타난다 '
    f"(DC+Fab −50%에서 풍력 {V('base', 'WT_GW'):.1f} → {V('DCFab_m50', 'WT_GW'):.1f} GW, PV {d('DCFab_m50', 'PV_GW'):+.1f} GW)."),
   ('4. MGA: ' if MGA else '4. MGA: 산정 실패', (
    f"비용을 B0보다 " + ', '.join(f"{float(nat.at[c, 'eps_pct']):g}% 더 허용하면 PV·풍력 총설비가 {-mga_dre(c):.1f} GW"
                                 f"({-100 * mga_dre(c) / V('B0', 'RE_GW'):.1f}%)" for c in MGA) + ' 줄어든다. '
    f"PV가 줄고({d(mL, 'PV_GW', 'B0'):+.1f} GW) 풍력이 늘며({d(mL, 'WT_GW', 'B0'):+.1f} GW), 출력제한률이 "
    f"{V('B0', 'curtail_%'):.1f}% → {V(mL, 'curtail_%'):.1f}%로 낮아진다. 같은 비용 허용률에서 2040년보다 감축 비율이 크다. "
    f"다만 늘어나는 풍력이 서울경기에 몰려 기술적 잠재량의 {pot_pct(mL, 'WT', 'GGD'):.0f}%가 된다.") if MGA else ''),
   ('5. 2040년과 비교: ', '2035년은 재생설비 규모가 작고 출력제한률도 낮아, 잠재량 상한의 설비 증가 효과는 작다. 수요 감축의 '
    'MWh당 회피비용은 2040년(약 15만 원)보다 다소 높다. CO2 잠재가격은 낮다(표 6-1).')])
T(summary_table(), font=7, caption='표 0-3. 전체 결과 요약 (MGA 비용 변화는 B0 대비, 나머지는 기준 대비)', fmts=FM)
F('fig1_overview.png', '그림 1. 케이스별 PV·풍력 설비(a)와 기준 대비 시스템 비용 변화(b)', width=17)

# 1장 기준
doc.add_page_break()
H('1. 2035년 기준 시나리오', 1)
F('fig0_base.png', '그림 2. 기준 발전량(a)·전력 소비(b) 구성과 지역별 PV·풍력 설비(c)', width=17)
B([('발전 구성: ', f"원자력 {V('base', 'nuclear_TWh'):.0f}, LNG {V('base', 'LNG_power_TWh'):.0f}, CHP {V('base', 'CHP_power_TWh'):.0f}, "
    f"PV {V('base', 'PV_TWh'):.0f}, 풍력 {V('base', 'WT_TWh'):.0f}, 연료전지 {V('base', 'fuelcell_power_TWh'):.0f}, "
    f"석탄 {V('base', 'coal_TWh'):.1f} TWh. CO2 상한 {V('base', 'CO2_cap_Mt'):.1f} Mt 안에서 석탄보다 배출계수가 낮은 LNG를 쓴다."),
   ('부문결합: ', f"수전해 {V('base', 'electrolyser_GW'):.1f} GW가 {V('base', 'electrolyser_elec_TWh'):.0f} TWh의 전력을 쓰고, "
    f"연료전지({V('base', 'fuelcell_GW'):.1f} GW)가 {V('base', 'fuelcell_power_TWh'):.0f} TWh를 되돌려준다. "
    f"ESS 신규는 {V('base', 'ESS_new_GWh'):.1f} GWh로 작다."),
   ('입지: ', f"신규 PV는 인천·대전세종충남·서울경기·광주전남·부산울산, 신규 풍력은 서울경기({R['base'].at['GGD', 'WT_GW']:.1f} GW)와 "
    f"충북({R['base'].at['CBD', 'WT_GW']:.1f} GW)에 집중된다. 잠재량 20%를 넘는 지역·기술은 "
    f"{len(set(moved_out))}개 지역에 걸쳐 있다(그림 2c 빨간 선).")])

# 2장 잠재량 상한
doc.add_page_break()
H('2. 지역별 잠재량 20% 상한', 1)
T(cap_table(), font=8, caption='표 2-1. 기준과 잠재량 20% 상한 비교', fmts=FM)
F('fig2_cap20.png', '그림 3. 지역별 PV(a)·풍력(b) 설비: 기준과 잠재량 20% 상한 (빨간 막대: 상한, ×: 기술적 잠재량)', width=17)
B([('비용: ', f"+{d('cap20', 'cost_sys'):.2f}조원/년(+{100 * d('cap20', 'cost_sys') / C0:.2f}%). PV·풍력 총량은 "
    f"{RE0:.1f} → {V('cap20', 'RE_GW'):.1f} GW로 거의 같다."),
   ('입지 재배치: ', f"상한에 걸리는 인천·부산울산 PV와 서울경기·충북 풍력이 줄고({cap_loss_txt} GW), 서울경기·충북·강원·경남 PV와 "
    f"여러 지역 풍력으로 옮겨간다({cap_gain_txt} GW)."),
   ('송전: ', f"HVDC 수송량 {V('base', 'HVDC_TWh'):.0f} → {V('cap20', 'HVDC_TWh'):.0f} TWh, 송전 혼잡 "
    f"{V('base', 'line_congestion_h'):,.0f} → {V('cap20', 'line_congestion_h'):,.0f} 선로·시간. 수요지 인근 입지가 막혀 송전 부담이 커진다."),
   ('가격: ', f"CO2 잠재가격 {V('base', 'CO2_price'):.0f} → {V('cap20', 'CO2_price'):.0f}천원/t, 부하가중 전력가격 "
    f"{V('base', 'EL_price_loadw'):.0f} → {V('cap20', 'EL_price_loadw'):.0f}천원/MWh."),
   ('2040년과 차이: ', f"2040년에는 상한으로 재생설비가 {float(r40c.at['B0', 'RE_GW'] - r40.at['B0', 'RE_GW']):+.1f} GW 늘었으나 "
    f"2035년은 {d('cap20', 'RE_GW'):+.1f} GW이다. 2035년은 보급 규모가 작아 상한 안에서 재배치할 여유가 크다.")])

# 3장 수요
doc.add_page_break()
H('3. 전력수요 감축 (일반·DC·Fab)', 1)
T(demand_table(), font=6.5, caption='표 3-1. 수요 감축 케이스별 결과', fmts=FM)
F('fig3_demand.png', '그림 4. 수요 감소량과 비용 감소(a), 케이스별 PV·풍력 변화(b)')
B([('회피비용: ', f"MWh당 {AV.min():.0f}~{AV.max():.0f}천원. 일반 −5%({AV['EL_m05']:.0f}), −10%({AV['EL_m10']:.0f}), "
    f"DC −25%({AV['DC_m25']:.0f}), Fab −25%({AV['Fab_m25']:.0f}), DC+Fab −25%({AV['DCFab_m25']:.0f}), −50%({AV['DCFab_m50']:.0f}). "
    '감축 대상보다 감축량이 결과를 결정한다.'),
   ('풍력이 먼저 줄어든다: ', f"일반 −10%에서 풍력 {d('EL_m10', 'WT_GW'):+.1f} GW, PV {d('EL_m10', 'PV_GW'):+.1f} GW. "
    f"DC+Fab −50%에서는 풍력 {d('DCFab_m50', 'WT_GW'):+.1f} GW로 신규 풍력이 {V('DCFab_m50', 'WT_new_GW'):.1f} GW만 남는다. "
    '서울경기 풍력이 가장 먼저 빠진다(그림 5).'),
   ('가스·석탄: ', f"CO2 상한이 구속되므로 가스·석탄 총배출은 같다. 수요가 줄면 석탄 발전({V('base', 'coal_TWh'):.1f} → "
    f"{V('DCFab_m50', 'coal_TWh'):.1f} TWh)이 줄고 그만큼 배출 여유로 LNG 발전이 약간 늘어난다"
    f"({V('base', 'LNG_power_TWh'):.0f} → {V('DCFab_m50', 'LNG_power_TWh'):.0f} TWh)."),
   ('출력제한: ', f"{V('base', 'curtail_%'):.1f}% → 일반 −10% {V('EL_m10', 'curtail_%'):.1f}%, DC+Fab −50% "
    f"{V('DCFab_m50', 'curtail_%'):.1f}%. 풍력이 먼저 줄어 PV 비중이 커지므로 출력제한률은 오히려 조금 오른다."),
   ('CO2 부담: ', f"CO2 잠재가격 {V('base', 'CO2_price'):.0f} → {V('DCFab_m50', 'CO2_price'):.0f}천원/t (DC+Fab −50%). "
    '같은 배출 목표를 더 낮은 한계비용으로 달성한다.')])
F('fig4_regional_change.png', '그림 5. 지역별 PV·풍력 설비 변화 (수요 감축은 기준 대비, MGA는 B0 대비)', width=16)

# 4장 MGA
doc.add_page_break()
H('4. 비용 증가를 제한한 최소 재생설비 대안 (MGA)', 1)
B([('방법: ', '2040년 분석과 같다. 1단계에서 C ≤ (1+ε)·C_B0 아래 PV·풍력 신규 용량 합계를 최소화(K*)하고, 2단계에서 신규 용량 ≤ K*(1+10⁻⁴) '
    '아래 비용을 다시 최소화한다. 열보조공급량은 지역별 B0 이하로 제한하고, PyPSA_GUI의 사용자 정의 제약(CHP 포함 CO2, DR, DC/Fab RE 조달)을 유지한다.'),
   ('수치 처리: ', 'MGA-1의 1단계는 기본 Barrier 설정(barrier.algorithm 3)에서 수렴 기준 10⁻⁶·10⁻⁵ 모두 "Barrier cannot determine '
    'infeasibility"로 해를 얻지 못하였다. Barrier 시작점 알고리즘을 1(infeasibility-estimate start)로 바꾸고 수렴 기준 10⁻⁵에서 '
    '1·2단계를 풀었다(10⁻⁶에서는 여전히 미수렴). 문제 정의(제약·목적함수)는 동일하고, 2단계 비용은 상한의 100.0%로 정상 종료되었다. '
    'MGA-2와 B0는 기본 설정(10⁻⁶)에서 정상 종료되었다.')])
if MGA:
    T(mga_table(), font=8, caption='표 4-1. MGA 결과 요약 (요청 양식)', fmts=FM)
    F('fig5_mga.png', '그림 6. 비용 증가 허용률별 재생설비 감소율(2035년·2040년 비교)(a)와 MGA 대안의 전력 수급 변화(b)', width=17)
    B([('감축 규모: ', ', '.join(f"{L(c)} {-mga_dre(c):.1f} GW(비용 +{V(c, 'cost_sys') - B0:.2f}조원)" for c in MGA)
        + f". 2040년 같은 허용률의 감소({red40_txt} GW)와 비교해 "
        '절대량은 비슷하거나 크고, 비율로는 더 크다.'),
       ('메커니즘: ', f"PV {d(mL, 'PV_GW', 'B0'):+.1f} GW, 풍력 {d(mL, 'WT_GW', 'B0'):+.1f} GW. PV 잉여가 줄어 출력제한이 "
        f"{V('B0', 'curtail_TWh'):.0f} → {V(mL, 'curtail_TWh'):.0f} TWh로, 수전해 전력이 {d(mL, 'electrolyser_elec_TWh', 'B0'):+.1f} TWh, "
        f"연료전지 발전이 {d(mL, 'fuelcell_power_TWh', 'B0'):+.1f} TWh 줄어든다. CHP 발전({d(mL, 'CHP_power_TWh', 'B0'):+.1f} TWh)이 "
        '늘어 열부문에서도 일부 조정된다. 2040년과 같은 "PV→풍력 전환 + 잉여 순환 감소" 구조이다.'),
       ('입지 주의: ', f"늘어나는 풍력은 대부분 서울경기(B0 {R['B0'].at['GGD', 'WT_GW']:.1f} → {R[mL].at['GGD', 'WT_GW']:.1f} GW, "
        f"잠재량 {ggd_pot:.1f} GW의 {pot_pct(mL, 'WT', 'GGD'):.0f}%)이다. 2040년과 마찬가지로 정책 수치로 쓰려면 잠재량 상한 조건의 MGA를 함께 볼 필요가 있다."),
       ('가격 지표 주의: ', f"CO2 잠재가격 {V('B0', 'CO2_price'):.0f} → {V(mL, 'CO2_price'):.0f}천원/t. 재생 용량을 인위적으로 묶은 데 따른 잠재가격이다.")])

# 5장 2040 비교 + 수급
doc.add_page_break()
H('5. 전력 수급 상세', 1)
T(balance_table(), font=6, caption='표 5-1. 케이스별 전력·열·가스 수급 (TWh)')
T(regional_table(), font=6, caption='표 5-2. 지역별 PV·풍력 설비 (GW)')

doc.add_page_break()
H('6. 2040년 결과와 비교', 1)
T(y2040_table(), font=8, caption='표 6-1. 2035년과 2040년 주요 결과 비교 (2040년은 현재 입력 기준 MGA 분석의 B0·잠재량 상한 B0)')
B([('시스템 비용 수준: ', f"2035년 기준 비용({C0:.1f}조원)이 2040년({float(r40.at['B0', 'cost_sys']):.1f}조원)보다 큰 것은 "
    f"가스 연료비({V('base', 'c_opex_gas'):.1f} vs {float(r40.at['B0', 'c_opex_gas']):.1f}조원)와 연료전지 투자비"
    f"({V('base', 'c_capex_Fuelcell'):.1f} vs {float(r40.at['B0', 'c_capex_Fuelcell']):.1f}조원, 설비 {V('base', 'fuelcell_GW'):.1f} vs "
    f"{float(r40.at['B0', 'fuelcell_GW']):.1f} GW) 차이가 크기 때문이다. 비용은 기존 설비 고정비를 뺀 값이므로 연도 간 절대 수준 비교보다 "
    '같은 연도 안의 케이스 간 차이로 해석해야 한다.'),
   ('보급 규모: ', f"2035년 PV·풍력 {RE0:.0f} GW는 2040년({b40r:.0f} GW)의 {100 * RE0 / b40r:.0f}%이다. 출력제한률도 낮다."),
   ('잠재량 상한: ', '두 연도 모두 상한 비용은 1% 이하이다. 2035년은 상한 안에서 재배치할 여유가 커서 설비 총량 증가가 거의 없다.'),
   ('수요 감축: ', f"2035년 회피비용(약 {AV.mean():.0f}천원/MWh)이 2040년 분석(153~159천원/MWh)보다 높다. 다만 2040년 수요 분석은 "
    'interface 갱신 전 입력과 목적함수(열보조 페널티 포함) 기준이라 직접 비교에는 주의가 필요하다. 2035년에는 감축분 대부분이 '
    '연간화 투자비가 PV의 약 4.6배인 풍력의 신규 설치 회피로 나타난다.'),
   ('MGA: ', (f"같은 비용 허용률에서 2035년의 재생설비 감소율이 2040년보다 크다({L(mL)}: {-100 * mga_dre(mL) / V('B0', 'RE_GW'):.1f}% vs "
    f"{100 * (1 - float(r40.at[mL, 'RE_GW']) / b40r):.1f}%). 2035년에는 PV 감축 폭({d(mL, 'PV_GW', 'B0'):+.1f} GW)이 커서 "
    f"출력제한이 {V('B0', 'curtail_%'):.1f}% → {V(mL, 'curtail_%'):.1f}%로 크게 줄고, 수전해·연료전지 순환도 함께 줄어든다.")
    if mL in r40.index else '')])

# 7장 시사점
H('7. 종합 및 시사점', 1)
B([('입지 제약: ', f"2035년 보급 규모에서는 잠재량 20% 상한을 두어도 비용 증가가 작다(+{100 * d('cap20', 'cost_sys') / C0:.1f}%). "
    '다만 송전 부담이 크게 늘므로, '
    '입지 분산 정책은 송전망 보강 계획과 함께 검토해야 한다.'),
   ('수요 전망: ', f"전력수요 10% 차이(약 {nat.at['EL_m10', 'd_demand']:.0f} TWh)는 PV·풍력 약 {RE0 - V('EL_m10', 'RE_GW'):.0f} GW, "
    f"시스템 비용 약 {nat.at['EL_m10', 'd_cost']:.1f}조원/년의 차이로 이어진다. 특히 풍력 보급 규모가 수요 전망에 민감하다."),
   ('보급 부담 완화: ', (f"비용을 1~2% 더 허용하면 PV·풍력을 {-mga_dre(MGA[0]):.0f}~{-mga_dre(mL):.0f} GW 줄일 수 있다. "
    '풍력 비중 확대가 주 수단이지만 수도권 풍력 입지가 전제이다.') if MGA else ''),
   ('한계: ', '① 단일연도(2035) 최적화로 2040년까지의 설비 경로는 반영하지 않았다. ② 잠재량 20%는 일률적 가정이다. '
    '③ MGA 대안의 가격 지표는 잠재가격이다. ④ 수요 감축 비용(효율 투자 등)은 포함하지 않았다.')])

# 부록
doc.add_page_break()
H('부록. 산출 파일과 실행 방법', 1)
T(pd.DataFrame([
    ['민감도 7개 + 기준', 'results/sens2035_*', 'run_2035_sensitivity.py'],
    ['MGA B0·2%', 'results/mga2035_*', 'run_mga_sweep.py --eps 1 2 --no-penalty'],
    ['MGA 1% (재계산)', 'results/mga_sweep_*', 'run_mga_sweep.py --eps 1 --no-penalty --b0-dir ... --barrier-alg 1 (convergetol 1e-5 재시도에서 수렴)'],
    ['집계', OUT, 'analyze_2035.py'],
    ['보고서·그림·표', OUT, 'make_2035_report.py'],
], columns=['내용', '결과 폴더', '실행 스크립트']), font=8)
B([('솔버: ', 'CPLEX Barrier(Crossover 생략), 수렴 기준 10⁻⁶. 민감도 8개 케이스는 모두 첫 시도에 정상 종료.'),
   ('수요 조정: ', '대상 부하의 p_set(정적값·시계열)에 배율을 곱함. 부하 형태·지역 비중 유지.'),
   ('회피비용: ', '(기준 시스템 비용 − 케이스 시스템 비용) ÷ (기준 전력수요 − 케이스 전력수요). 시스템 비용은 열보조공급 페널티 제외.')])

path = os.path.join(OUT, '2035_민감도분석_보고서.docx')
doc.save(path)
print(f'완료: {path}')
pd.set_option('display.width', 250)
print(summary_table().round(2).to_string(index=False))
print(demand_table().round(1).to_string(index=False))
print(mga_table().round(2).to_string(index=False))
print(y2040_table().to_string(index=False))
print('moved_out', moved_out)
