"""
잠재량 20% 상한 조건 MGA 분석 보고서(Word) + 그림 + 표 생성 (무상한 MGA와 비교)

선행:
  python analyze_mga.py results/mga_sweep_cap20_<ts> --out reports/mga_cap20_<날짜>
  (무상한 결과: reports/mga_<날짜>/ — make_mga_report.py 입력과 동일)
사용법:
  python make_mga_cap_report.py [--cap-dir reports/mga_cap20_...] [--base-dir reports/mga_...]
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
ap.add_argument('--cap-dir', default=sorted(glob.glob('reports/mga_cap20_*'))[-1])
ap.add_argument('--base-dir', default=sorted(d for d in glob.glob('reports/mga_2*') if os.path.isdir(d))[-1])
ap.add_argument('--share', type=float, default=0.2)
args = ap.parse_args()
OUT = args.cap_dir
SHARE = args.share

cap = pd.read_csv(f'{args.cap_dir}/mga_national.csv', index_col=0)
capR = pd.read_csv(f'{args.cap_dir}/mga_regional.csv')
unc = pd.read_csv(f'{args.base_dir}/mga_national.csv', index_col=0)
uncR = pd.read_csv(f'{args.base_dir}/mga_regional.csv')

REGION_KR = {'BSN': '부산울산', 'CBD': '충북', 'CND': '대전세종충남', 'GBD': '대구경북', 'GGD': '서울경기',
             'GND': '경남', 'GWD': '강원', 'ICN': '인천', 'JBD': '전북', 'JJD': '제주', 'JND': '광주전남'}
REG = ['GGD', 'ICN', 'GWD', 'CBD', 'CND', 'JBD', 'GBD', 'BSN', 'GND', 'JND', 'JJD']
C_PV, C_WT, COL_B0, C_CAP, C_UNC, C_PEN, C_CHP, C_GAS = ('#F2B134', '#3C8DBC', '#7F7F7F', '#1F4E79', '#A6A6A6',
                                                          '#C0392B', '#8E5EA2', '#E07B39')


def eps_of(c):
    return float(c.split('-', 1)[1])


MGA = sorted([c for c in cap.index if c.startswith('MGA-')], key=eps_of)
CASES = ['B0'] + MGA
EPS = {c: eps_of(c) for c in MGA}
UMGA = sorted([c for c in unc.index if c.startswith('MGA-')], key=eps_of)
COMMON = [c for c in MGA if c in unc.index]

C_B0, RE_B0 = float(cap.at['B0', 'cost_sys']), float(cap.at['B0', 'RE_GW'])
U_B0, URE_B0 = float(unc.at['B0', 'cost_sys']), float(unc.at['B0', 'RE_GW'])

LIM = {t: {r: max(SHARE * TECH_POTENTIAL_GW[t][r], float(capR[(capR.case == 'B0') & (capR.region == r)][f'{t}_exist_GW'].iloc[0]))
           for r in REG} for t in ('PV', 'WT')}


def V(c, k, df=cap):
    return float(df.at[c, k])


def U(c, k):
    return float(unc.at[c, k])


def inc(c, df=cap):
    base = C_B0 if df is cap else U_B0
    return 100 * (V(c, 'cost_sys', df) / base - 1)


def dre(c, df=cap):
    return V(c, 'RE_GW', df) - (RE_B0 if df is cap else URE_B0)


def g(c, k):
    return V(c, k) - V('B0', k)


def gu(c, k):
    return U(c, k) - U('B0', k)


R = {c: capR[capR.case == c].set_index('region').reindex(REG) for c in CASES}
RU = {c: uncR[uncR.case == c].set_index('region').reindex(REG) for c in ['B0'] + UMGA}


def binding(c, t):
    return [r for r in REG if R[c].at[r, f'{t}_GW'] >= LIM[t][r] - 0.05]


def headroom(c, t):
    return sum(max(LIM[t][r] - R[c].at[r, f'{t}_GW'], 0.0) for r in REG)


figs = {}


def save(fig, name, rect=(0, 0, 1, 1)):
    p = os.path.join(OUT, name)
    fig.tight_layout(rect=rect)
    fig.savefig(p, dpi=200)
    plt.close(fig)
    figs[name] = p


# 구간별 한계비용 (같은 ε 구간끼리 비교, 무상한은 MGA-0 건너뜀)
prev = ['B0'] + MGA[:-1]
SEG = [f"{(EPS[p] if p != 'B0' else 0):g}→{EPS[c]:g}%" for p, c in zip(prev, MGA)]


def seg_mc(df, cases):
    pv = ['B0'] + cases[:-1]
    mc = [(V(c, 'cost_sys', df) - V(p, 'cost_sys', df)) * 1e4 / max(V(p, 'RE_GW', df) - V(c, 'RE_GW', df), 1e-9)
          for p, c in zip(pv, cases)]
    rd = [V(p, 'RE_GW', df) - V(c, 'RE_GW', df) for p, c in zip(pv, cases)]
    return mc, rd


mc_cap, red_cap = seg_mc(cap, MGA)
mc_unc, red_unc = seg_mc(unc, COMMON) if len(COMMON) == len(MGA) else ([np.nan] * len(MGA), [np.nan] * len(MGA))

# ───────────── 그림 ─────────────
# fig1 경계 비교: (a) 자기 B0 대비 증가율, (b) 절대 비용, (c) 구간별 한계비용
fig, ax = plt.subplots(1, 3, figsize=(14, 3.9))
ux = [0.0] + [inc(c, unc) for c in UMGA]; uy = [URE_B0] + [U(c, 'RE_GW') for c in UMGA]
cx = [0.0] + [inc(c) for c in MGA]; cy = [RE_B0] + [V(c, 'RE_GW') for c in MGA]
ax[0].plot(ux, uy, '--o', color=C_UNC, ms=4, label='상한 해제 (기존 분석)')
ax[0].plot(cx, cy, '-o', color=C_CAP, label=f'잠재량 {SHARE:.0%} 상한')
for c in MGA:
    ax[0].annotate(f'{dre(c):+.1f}', (inc(c), V(c, 'RE_GW')), textcoords='offset points', xytext=(4, 5), fontsize=7.5,
                   color=C_CAP)
ax[0].set_xlabel('비용 증가율 (%, 각 조건의 B0 대비)'); ax[0].set_ylabel('PV + 풍력 총 설비 (GW)')
ax[0].set_title('(a) 비용 증가 허용률별 최소 재생설비'); ax[0].grid(alpha=0.3); ax[0].legend(fontsize=7.5)

ax[1].plot([U_B0] + [U(c, 'cost_sys') for c in UMGA], uy, '--o', color=C_UNC, ms=4, label='상한 해제')
ax[1].plot([C_B0] + [V(c, 'cost_sys') for c in MGA], cy, '-o', color=C_CAP, label=f'잠재량 {SHARE:.0%} 상한')
ax[1].scatter([U_B0], [URE_B0], s=60, color=COL_B0, zorder=4)
ax[1].scatter([C_B0], [RE_B0], s=60, color=C_CAP, zorder=4)
ax[1].annotate('상한 해제 B0', (U_B0, URE_B0), textcoords='offset points', xytext=(4, -12), fontsize=7.5)
ax[1].annotate(f'상한 B0 (+{C_B0 - U_B0:.2f}조)', (C_B0, RE_B0), textcoords='offset points', xytext=(4, 4), fontsize=7.5)
ax[1].set_ylim(top=max(cy + uy) + 10)
ax[1].set_xlabel('시스템 비용 (조원/년)'); ax[1].set_ylabel('PV + 풍력 총 설비 (GW)')
ax[1].set_title('(b) 절대 비용 기준 비교'); ax[1].grid(alpha=0.3); ax[1].legend(fontsize=7.5)

xx = np.arange(len(MGA))
ax[2].bar(xx - 0.2, mc_unc, 0.4, color=C_UNC, label='상한 해제')
ax[2].bar(xx + 0.2, mc_cap, 0.4, color=C_CAP, label=f'잠재량 {SHARE:.0%} 상한')
for i, (m_, r_) in enumerate(zip(mc_cap, red_cap)):
    ax[2].text(i + 0.2, m_ + 25, f'{m_:,.0f}\n(-{r_:.1f})', ha='center', fontsize=6.8, color=C_CAP)
ax[2].set_xticks(xx, SEG, fontsize=7.5); ax[2].set_ylim(0, max(mc_cap) * 1.3)
ax[2].set_ylabel('억원/년 per GW 감축'); ax[2].set_title('(c) 구간별 1 GW 감축당 추가비용')
ax[2].grid(axis='y', alpha=0.3); ax[2].legend(fontsize=7.5, loc='upper left')
save(fig, 'fig1_frontier_compare.png')

# fig2 설비 구성과 출력제한
fig, ax = plt.subplots(figsize=(9, 3.6))
xx = np.arange(len(CASES))
pv = [V(c, 'PV_GW') for c in CASES]; wt = [V(c, 'WT_GW') for c in CASES]
ax.bar(xx, pv, color=C_PV, label='PV', width=0.6)
ax.bar(xx, wt, bottom=pv, color=C_WT, label='풍력', width=0.6)
for i, c in enumerate(CASES):
    ax.text(i, pv[i] / 2, f'{pv[i]:.0f}', ha='center', fontsize=7.5)
    ax.text(i, pv[i] + wt[i] / 2, f'{wt[i]:.0f}', ha='center', fontsize=7.5, color='white')
    ax.text(i, pv[i] + wt[i] + 4, f'{pv[i] + wt[i]:.0f}', ha='center', fontsize=8, fontweight='bold')
um = [U(c, 'RE_GW') if c in unc.index else np.nan for c in CASES]
ax.scatter(xx + 0.36, um, marker='_', s=200, color='k', zorder=4, label='상한 해제 총설비')
ax.set_xticks(xx, CASES); ax.set_ylabel('GW'); ax.set_ylim(0, 460)
ax2 = ax.twinx()
ax2.plot(xx, [V(c, 'curtail_%') for c in CASES], '-o', color=C_PEN, label='출력제한률')
ax2.set_ylabel('출력제한률 (%)'); ax2.set_ylim(0, 15)
h1, l1 = ax.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
ax.legend(h1 + h2, l1 + l2, fontsize=7.5, loc='upper right', ncol=4)
ax.set_title(f'잠재량 {SHARE:.0%} 상한: 케이스별 PV·풍력 설비와 출력제한률'); ax.grid(axis='y', alpha=0.3)
save(fig, 'fig2_capacity_mix.png')

# fig3 지역별 설비와 상한
SHOW = [c for c in ('B0', 'MGA-1', 'MGA-3') if c in CASES]
fig, ax = plt.subplots(1, 2, figsize=(12, 4.2), sharey=True)
yy = np.arange(len(REG))[::-1]
cols = ['#BDD7EE', '#2E75B6', '#1F4E79']
for k, t in enumerate(('PV', 'WT')):
    for j, c in enumerate(SHOW):
        ax[k].barh(yy + 0.27 - 0.27 * j, R[c][f'{t}_GW'], 0.27, color=cols[j], label=c)
    ax[k].scatter([LIM[t][r] for r in REG], yy, marker='|', s=260, color=C_PEN, zorder=4, label=f'상한 (잠재량 {SHARE:.0%})')
    ax[k].scatter(RU['B0'][f'{t}_GW'], yy, marker='x', s=22, color='k', zorder=4, label='상한 해제 B0')
    ax[k].set_title(f"({'ab'[k]}) {'태양광' if t == 'PV' else '풍력'} 지역별 설비 (GW)")
    ax[k].grid(axis='x', alpha=0.3)
ax[0].set_yticks(yy, [REGION_KR[r] for r in REG])
h, l = ax[0].get_legend_handles_labels()
fig.legend(h, l, fontsize=8, loc='lower center', ncol=5, bbox_to_anchor=(0.5, -0.01))
save(fig, 'fig3_regional.png', rect=(0, 0.06, 1, 1))

# fig4 메커니즘
items = [('PV 발전', 'PV_TWh'), ('풍력 발전', 'WT_TWh'), ('출력제한', 'curtail_TWh'),
         ('LNG 발전', 'LNG_power_TWh'), ('CHP 발전', 'CHP_power_TWh'),
         ('히트펌프 전력', 'HP_elec_TWh'), ('수전해 전력', 'electrolyser_elec_TWh'),
         ('EV 충전(V2G 포함)', 'EV_charge_TWh'), ('ESS 충전', 'ESS_charge_TWh')]
SER = [(f'상한 {c}', [g(c, k) for _, k in items], col, None) for c, col in zip(('MGA-1', 'MGA-3'), ('#2E75B6', '#1F4E79'))]
SER += [(f'상한 해제 {c}', [gu(c, k) for _, k in items], 'white', '///') for c in ('MGA-1', 'MGA-3') if c in unc.index]
fig, ax = plt.subplots(figsize=(11.5, 3.9))
xx = np.arange(len(items)); wdt = 0.8 / len(SER)
for j, (lb, vals, col, hat) in enumerate(SER):
    ax.bar(xx - 0.4 + wdt * (j + 0.5), vals, wdt, label=lb, color=col, hatch=hat,
           edgecolor='#555555' if hat else col, lw=0.6)
ax.axhline(0, color='k', lw=0.6)
ax.axvspan(-0.5, 2.5, color=C_PV, alpha=0.08); ax.axvspan(2.5, 4.5, color=C_GAS, alpha=0.08)
ax.axvspan(4.5, 8.5, color=C_WT, alpha=0.08)
ax.set_xticks(xx, [n_ for n_, _ in items], fontsize=8)
ax.set_ylabel('각 조건 B0 대비 변화 (TWh)'); ax.legend(fontsize=7.5, loc='lower left', ncol=2); ax.grid(axis='y', alpha=0.3)
ax.set_title('재생설비 감축 메커니즘: 각 조건의 B0 대비 연간 전력 수급 변화')
save(fig, 'fig4_mechanism.png')

# fig5 비용 구성 변화
groups = [('PV 투자비', ['c_capex_PV']), ('풍력 투자비', ['c_capex_WT']),
          ('가스 연료비(발전기)', ['c_opex_gas']), ('CHP 운전비', ['c_opex_CHP']),
          ('V2G·ESS·연료전지 운전비', ['c_opex_EV_V2G', 'c_opex_ESS_Ch', 'c_opex_ESS_Disch', 'c_opex_Fuelcell']),
          ('수소·열저장 투자비', ['c_capex_Electrolyser', 'c_capex_Fuelcell', 'c_capex_H2_Storage', 'c_capex_H_Storage'])]
colors = [C_PV, C_WT, C_GAS, C_CHP, '#2E7D32', '#9E9E9E']


def cdiff(c, k):
    a = cap.at[c, k] if k in cap else np.nan
    b = cap.at['B0', k] if k in cap else np.nan
    return (0 if np.isnan(a) else a) - (0 if np.isnan(b) else b)


fig, ax = plt.subplots(figsize=(10.5, 3.8))
xx = np.arange(len(MGA))
pos = np.zeros(len(MGA)); neg = np.zeros(len(MGA))
for (nm, ks), col in zip(groups, colors):
    vals = np.array([sum(cdiff(c, k) for k in ks) for c in MGA])
    base = np.where(vals >= 0, pos, neg)
    ax.bar(xx, vals, 0.6, bottom=base, color=col, label=nm)
    pos += np.where(vals >= 0, vals, 0); neg += np.where(vals < 0, vals, 0)
tot = [V(c, 'cost_sys') - C_B0 for c in MGA]
ax.plot(xx, tot, 'k_', markersize=26, mew=2.5, label='순 비용 증가')
for i, t in enumerate(tot):
    ax.text(i + 0.33, t, f'+{t:.2f}', va='center', fontsize=8)
ax.axhline(0, color='k', lw=0.6)
ax.set_xticks(xx, MGA); ax.set_ylabel('B0 대비 변화 (조원/년)'); ax.grid(axis='y', alpha=0.3)
ax.legend(fontsize=7.5, loc='center left', bbox_to_anchor=(1.01, 0.5)); ax.set_title(f'잠재량 {SHARE:.0%} 상한: 비용 구성 변화 (B0 대비)')
save(fig, 'fig5_cost_breakdown.png')


# ───────────── 표 ─────────────
def main_table():
    rows = []
    for c in CASES:
        rows.append({'시나리오': c, '비용 증가 허용률': '—' if c == 'B0' else f'{EPS[c]:g}%',
                     '시스템 비용 (조원)': V(c, 'cost_sys'), '실제 비용 증가율 (%)': 0.0 if c == 'B0' else inc(c),
                     '태양광 GW': V(c, 'PV_GW'), '풍력 GW': V(c, 'WT_GW'), 'B0 대비 재생 감소 GW': -dre(c),
                     'ESS GW/GWh': f"{V(c, 'ESS_GW'):.1f} / {V(c, 'ESS_GWh'):.0f}", '출력제한율 (%)': V(c, 'curtail_%')})
    return pd.DataFrame(rows)


def detail_table():
    rows = []
    for c in CASES:
        rows.append({'시나리오': c, '비용 상한 (조원)': '—' if c == 'B0' else f"{cap.at[c, 'cost_cap'] / 1e9:.3f}",
                     '1단계 최소 신규 K* (GW)': V(c, 'stage1_K_MW') / 1e3,
                     '최종 PV 신규 (GW)': V(c, 'PV_new_GW'), '최종 풍력 신규 (GW)': V(c, 'WT_new_GW'),
                     '최종 신규 합계 (GW)': V(c, 'RE_new_GW'),
                     '2단계 비용 (조원)': V(c, 'cost_sys'), '배출량 (Mt)': V(c, 'CO2_Mt'),
                     '열보조공급 (TWh)': V(c, 'heat_fallback_TWh'),
                     '풍력 남은 여유 (GW)': headroom(c, 'WT'), 'PV 남은 여유 (GW)': headroom(c, 'PV')})
    return pd.DataFrame(rows)


def compare_table():
    rows = []
    for c in ['B0'] + COMMON:
        rows.append({'ε': '—' if c == 'B0' else f'{EPS[c]:g}%',
                     '해제: 비용 (조원)': U(c, 'cost_sys'), '상한: 비용 (조원)': V(c, 'cost_sys'),
                     '해제: PV/풍력 (GW)': f"{U(c, 'PV_GW'):.1f} / {U(c, 'WT_GW'):.1f}",
                     '상한: PV/풍력 (GW)': f"{V(c, 'PV_GW'):.1f} / {V(c, 'WT_GW'):.1f}",
                     '해제: 감소 (GW)': 0.0 if c == 'B0' else -dre(c, unc),
                     '상한: 감소 (GW)': 0.0 if c == 'B0' else -dre(c),
                     '해제: 출력제한 (%)': U(c, 'curtail_%'), '상한: 출력제한 (%)': V(c, 'curtail_%'),
                     '해제: CHP 연료 (TWh)': U(c, 'CHP_fuel_TWh'), '상한: CHP 연료 (TWh)': V(c, 'CHP_fuel_TWh')})
    return pd.DataFrame(rows)


def b0_table():
    items = [('시스템 비용 (조원)', 'cost_sys'), ('PV (GW)', 'PV_GW'), ('풍력 (GW)', 'WT_GW'), ('PV+풍력 (GW)', 'RE_GW'),
             ('출력제한율 (%)', 'curtail_%'), ('PV 이용률 (%)', 'PV_CF_%'), ('풍력 이용률 (%)', 'WT_CF_%'),
             ('PV 투자비 (조원)', 'c_capex_PV'), ('풍력 투자비 (조원)', 'c_capex_WT'), ('가스 연료비 (조원)', 'c_opex_gas'),
             ('수전해 전력 (TWh)', 'electrolyser_elec_TWh'), ('HVDC 수송량 (TWh)', 'HVDC_TWh'),
             ('CO2 잠재가격 (천원/t)', 'CO2_price'), ('부하가중 전력가격 (천원/MWh)', 'EL_price_loadw')]
    return pd.DataFrame([{'항목': a, '상한 해제 B0': U('B0', k), f'잠재량 {SHARE:.0%} 상한 B0': V('B0', k),
                          '차이': V('B0', k) - U('B0', k)} for a, k in items])


def balance_table():
    items = [('PV 발전량', 'PV_TWh'), ('풍력 발전량', 'WT_TWh'), ('출력제한량', 'curtail_TWh'),
             ('원자력', 'nuclear_TWh'), ('LNG 발전', 'LNG_power_TWh'), ('CHP 발전', 'CHP_power_TWh'),
             ('CHP 열', 'CHP_heat_TWh'), ('CHP 연료 (가스)', 'CHP_fuel_TWh'), ('가스 연료 총량', 'gas_fuel_TWh'),
             ('히트펌프 전력', 'HP_elec_TWh'), ('히트펌프 열', 'HP_heat_TWh'), ('수전해 전력', 'electrolyser_elec_TWh'),
             ('수전해 설비 (GW)', 'electrolyser_GW'), ('연료전지 발전', 'fuelcell_power_TWh'),
             ('EV 충전 (V2G 포함)', 'EV_charge_TWh'), ('V2G 방전', 'V2G_TWh'), ('ESS 충전', 'ESS_charge_TWh'),
             ('전력 소비 합계', 'elec_demand_total_TWh'), ('HVDC 수송량', 'HVDC_TWh'),
             ('CO2 잠재가격 (천원/t)', 'CO2_price'), ('부하가중 전력가격 (천원/MWh)', 'EL_price_loadw')]
    df = pd.DataFrame({'항목 (TWh)': [a for a, _ in items]})
    for c in CASES:
        df[c] = [V(c, k) for _, k in items]
    return df


def regional_table():
    rows = []
    for r in REG:
        row = {'지역': REGION_KR[r], 'PV 상한': LIM['PV'][r]}
        for c in SHOW:
            row[f'{c} PV'] = R[c].at[r, 'PV_GW']
        row['해제 B0 PV'] = RU['B0'].at[r, 'PV_GW']
        row['풍력 상한'] = LIM['WT'][r]
        for c in SHOW:
            row[f'{c} 풍력'] = R[c].at[r, 'WT_GW']
        row['해제 B0 풍력'] = RU['B0'].at[r, 'WT_GW']
        rows.append(row)
    return pd.DataFrame(rows)


mc_tbl = pd.DataFrame({'구간': SEG, '상한: 추가비용 (조원/년)': [V(c, 'cost_sys') - V(p, 'cost_sys') for p, c in zip(prev, MGA)],
                       '상한: 재생 감소 (GW)': red_cap, '상한: 1 GW당 (억원/년)': mc_cap,
                       '해제: 재생 감소 (GW)': red_unc, '해제: 1 GW당 (억원/년)': mc_unc})

TBL = {'표1_결과요약': main_table(), '표2_단계별상세': detail_table(), '표3_상한해제비교': compare_table(),
       '표4_구간효율': mc_tbl, '표5_B0비교': b0_table(), '표6_수급변화': balance_table(), '표7_지역별설비': regional_table()}
with pd.ExcelWriter(os.path.join(OUT, 'MGA_잠재량상한_분석_표.xlsx')) as xw:
    for k, v in TBL.items():
        v.to_excel(xw, sheet_name=k, index=False)
    cap.T.to_excel(xw, sheet_name='전국_전체지표')
    capR.to_excel(xw, sheet_name='지역_전체지표', index=False)

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


def T(df, font=8, caption=None, fmt='{:,.2f}'):
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
            txt = fmt.format(v) if isinstance(v, (float, np.floating)) else str(v)
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


m01, m05, m1, m2, m3 = 'MGA-0.1', 'MGA-0.5', 'MGA-1', 'MGA-2', 'MGA-3'
WT_LIM = sum(LIM['WT'].values()); PV_LIM = sum(LIM['PV'].values())
pv_cut = {r: R['B0'].at[r, 'PV_GW'] - R[m3].at[r, 'PV_GW'] for r in REG}
pv_cut_main = [r for r in sorted(pv_cut, key=pv_cut.get, reverse=True) if pv_cut[r] > 0.5]
pv_cut_text = ', '.join(f"{REGION_KR[r]} {R['B0'].at[r, 'PV_GW']:.1f} → {R[m3].at[r, 'PV_GW']:.1f}" for r in pv_cut_main)
wt_add = {r: R[m3].at[r, 'WT_GW'] - R['B0'].at[r, 'WT_GW'] for r in REG}
wt_add_main = [r for r in sorted(wt_add, key=wt_add.get, reverse=True) if wt_add[r] > 0.1]
same_cost_unc = float(np.interp(V(m1, 'cost_sys'), [U_B0] + [U(c, 'cost_sys') for c in UMGA], uy))
re_unc_at_capB0 = float(np.interp(C_B0, [U_B0] + [U(c, 'cost_sys') for c in UMGA], uy))

doc.add_heading(f'잠재량 {SHARE:.0%} 상한 조건의 최소 재생설비 대안 분석 (MGA)', 0)
P('2040년 재생에너지 보급 상한 해제 시나리오 + 지역별 태양광·풍력 기술적 잠재량 20% 상한 — '
  f"비용 증가 허용률 {', '.join(f'{EPS[c]:g}' for c in MGA)}%", size=12)
P(f'작성일: {datetime.now():%Y-%m-%d}   |   모형: PyPSA 0.31.1 + linopy + CPLEX Barrier, 8,784시간 단일연도   |   '
  '비교 대상: 상한 해제 MGA 분석(기존 보고서)', size=9)

H('개요', 1)
P('기존 MGA 분석에서는 늘어나는 풍력이 서울경기에 몰려 기술적 잠재량을 넘었다. 이번 분석은 지역별 PV·풍력 설비를 '
  f'max(기술적 잠재량 × {SHARE:.0%}, 기존 설비)로 제한한 뒤, 같은 방법으로 B0와 MGA를 다시 계산하였다. '
  '상한 조건이 바뀌면 비용 최적해도 바뀌므로 B0를 새로 구하고, 모든 MGA 케이스는 이 새 B0 비용을 기준으로 하였다.')
T(pd.DataFrame([
    ['기준 입력', '현재 interface (2040년, 재생 보급 상한 해제). 입력 파일은 변경하지 않음'],
    ['추가 조건', f'지역별 PV·풍력 p_nom_max = max(기술적 잠재량 × {SHARE:.0%}, 기존 설비) — 메모리에서만 적용. '
     f'전국 상한 합계 PV {PV_LIM:.0f} GW, 풍력 {WT_LIM:.0f} GW'],
    ['B0', f'상한 조건에서 비용 최소화 (전력 slack 0). C_B0 = {C_B0:.4f}조원/년'],
    ['MGA-ε', f"ε = {', '.join(f'{EPS[c]:g}' for c in MGA)}% — 모두 상한 조건 B0 비용 기준으로 별도 실행"],
    ['방법', '기존 분석과 동일: 1단계 C ≤ (1+ε)·C_B0 하에서 PV·풍력 신규 용량 최소화 → 2단계 신규 용량 ≤ K*(1+10⁻⁴) 하에서 비용 최소화. '
     '비용 정의(목적함수 − 열보조공급 페널티), 열보조공급량 B0 이하 제한, CO2·CHP·DR·DC/Fab 제약 모두 동일'],
], columns=['항목', '내용']), font=8, caption='표 0-1. 분석 구성')

H('핵심 결과', 2)
B([('1. 잠재량 상한 자체의 비용은 약 1%이고, 상한 조건 B0는 재생설비가 오히려 많다. ',
    f"B0 비용이 {U_B0:.2f} → {C_B0:.2f}조원/년(+{100 * (C_B0 / U_B0 - 1):.2f}%)으로 오른다. 풍력 총량은 "
    f"{U('B0', 'WT_GW'):.1f} → {V('B0', 'WT_GW'):.1f} GW로 거의 같고 입지만 서울경기에서 다른 지역으로 흩어진다. "
    f"대신 PV는 {U('B0', 'PV_GW'):.1f} → {V('B0', 'PV_GW'):.1f} GW로 늘어, 총설비는 {URE_B0:.1f} → {RE_B0:.1f} GW이다. "
    f"수요지 인근 입지가 막혀 설비가 먼 지역으로 흩어지면서 송전 부담(HVDC 수송량 {U('B0', 'HVDC_TWh'):.0f} → "
    f"{V('B0', 'HVDC_TWh'):.0f} TWh)과 출력제한({U('B0', 'curtail_%'):.1f} → {V('B0', 'curtail_%'):.1f}%)이 늘기 때문이다."),
   ('2. 비용 0.5~1% 구간까지는 상한 해제와 거의 같은 비율로 줄일 수 있다. ',
    f"상한 조건 B0 대비 0.1%에 {-dre(m01):.1f} GW, 0.5%에 {-dre(m05):.1f} GW, 1%에 {-dre(m1):.1f} GW가 줄어든다"
    f"(상한 해제: {-dre(m01, unc):.1f}, {-dre(m05, unc):.1f}, {-dre(m1, unc):.1f} GW)."),
   ('3. 2% 이후에는 감축 여지가 빠르게 사라진다. ',
    f"2%에 {-dre(m2):.1f} GW, 3%에 {-dre(m3):.1f} GW로, 상한 해제({-dre(m2, unc):.1f}, {-dre(m3, unc):.1f} GW)보다 "
    f"각각 {dre(m2) - dre(m2, unc):.1f}, {dre(m3) - dre(m3, unc):.1f} GW 적다. 2→3% 구간의 1 GW 감축당 비용은 "
    f"{mc_cap[-1]:,.0f}억원/년으로 상한 해제({mc_unc[-1]:,.0f}억원/년)의 {mc_cap[-1] / mc_unc[-1]:.1f}배이다. "
    f"풍력을 더 지을 곳이 거의 없기 때문이다(MGA-3의 풍력 남은 여유 {headroom(m3, 'WT'):.1f} GW, 대부분 제주)."),
   ('4. 풍력 증가는 대구경북 한 곳뿐이다. ',
    f"B0에서 이미 풍력 11개 지역 중 {len(binding('B0', 'WT'))}곳이 상한에 걸려 있다. MGA의 풍력 증가는 대구경북 "
    f"{R['B0'].at['GBD', 'WT_GW']:.1f} → {R[m2].at['GBD', 'WT_GW']:.1f} GW(MGA-2에서 상한 {LIM['WT']['GBD']:.1f} GW 도달)가 "
    f"거의 전부이다. 전국 풍력 증가는 {g(m3, 'WT_GW'):+.1f} GW로 상한 해제 MGA-3({gu(m3, 'WT_GW'):+.1f} GW)의 절반 수준이다."),
   ('5. PV 감축은 상한에 걸리지 않은 지역에서 일어난다. ',
    f"MGA-3 기준 {', '.join(f'{REGION_KR[r]} {-pv_cut[r]:+.1f}' for r in pv_cut_main)} GW. 상한에 붙어 있는 "
    f"지역({'·'.join(REGION_KR[r] for r in binding(m3, 'PV'))})의 PV는 MGA에서도 상한을 유지하고, "
    '상한 해제 B0에서는 기존 설비만 있던(즉 비용상 덜 유리한) 지역에 상한 때문에 밀려 들어간 PV부터 빠진다.'),
   ('6. 풍력 대신 열부문 전환이 더 일찍, 더 크게 나타난다. ',
    f"CHP 연료가 MGA-1에서 이미 {V('B0', 'CHP_fuel_TWh'):.1f} → {V(m1, 'CHP_fuel_TWh'):.1f} TWh로 늘어 히트펌프 전력이 "
    f"{g(m1, 'HP_elec_TWh'):+.1f} TWh 줄어든다(상한 해제에서는 3%부터 나타나던 현상). 2% 이상에서는 수전해 설비"
    f"({V('B0', 'electrolyser_GW'):.1f} → {V(m3, 'electrolyser_GW'):.1f} GW)와 열저장 투자가 늘어 잉여 PV 흡수를 늘린다. "
    '신규 ESS는 여전히 0이다.'),
   ('7. 같은 절대 비용에서 비교하면 상한의 영향이 분명하다. ',
    f"상한 조건 MGA-1의 비용({V(m1, 'cost_sys'):.2f}조원)이면 상한 해제 조건에서는 총설비를 {same_cost_unc:.1f} GW까지 줄일 수 있었으나, "
    f"상한 조건에서는 {V(m1, 'RE_GW'):.1f} GW가 최소이다. 상한 B0 비용({C_B0:.2f}조원)으로 상한 해제 조건에서는 "
    f"{re_unc_at_capB0:.1f} GW까지 가능하다.")])
T(main_table(), font=8, caption='표 0-2. 결과 요약 (요청 양식, 잠재량 20% 상한 조건)')
F('fig1_frontier_compare.png', '그림 1. 상한 해제와 잠재량 20% 상한의 최소 재생설비 경계: (a) 각 조건 B0 대비 증가율, '
  '(b) 절대 비용, (c) 구간별 1 GW 감축당 추가비용', width=17)

# 1장 방법
doc.add_page_break()
H('1. 분석 방법과 B0 변화', 1)
H('1.1 방법 (기존 분석과 동일)', 2)
B([('잠재량 상한: ', f"지역별 PV·풍력 확장 상한을 max(기술적 잠재량 × {SHARE:.0%}, 기존 설비)로 둔다. 기존 설비가 20%를 넘는 지역"
    '(예: 부산울산·인천 풍력)은 기존 설비가 상한이 된다. 입력 파일과 PyPSA_GUI.py는 수정하지 않고 네트워크 생성 후 메모리에서만 적용하였다.'),
   ('B0 재계산: ', '상한 조건에서 비용 최소해가 달라지므로 B0를 다시 계산하고, 모든 MGA 케이스는 이 B0 비용을 기준으로 하였다.'),
   ('나머지: ', '1·2단계 절차, 비용 정의, 전력 slack 0, 열보조공급량 B0 이하 제한, 사용자 정의 제약(CHP 포함 CO2, DR, DC/Fab) 유지 등은 '
    '기존 상한 해제 MGA와 같다. 모든 케이스가 정상(optimal) 종료되었고, 배출량은 상한(58.87 Mt)과 같다.')])
T(detail_table(), font=7.5, caption='표 1-1. 단계별 상세 결과 (풍력·PV 남은 여유 = 지역별 상한까지 남은 용량의 합)')
H('1.2 잠재량 상한이 B0에 주는 영향', 2)
T(b0_table(), font=8, caption='표 1-2. 상한 해제 B0와 잠재량 20% 상한 B0 비교')
B([('입지 재배치: ', f"상한 해제 B0에서 잠재량 20%를 넘던 설비는 서울경기 풍력({RU['B0'].at['GGD', 'WT_GW']:.1f} GW, 상한 "
    f"{LIM['WT']['GGD']:.1f}), 인천 PV({RU['B0'].at['ICN', 'PV_GW']:.1f}, 상한 {LIM['PV']['ICN']:.1f}), 부산울산 PV"
    f"({RU['B0'].at['BSN', 'PV_GW']:.1f}, 상한 {LIM['PV']['BSN']:.1f}), 광주전남 PV({RU['B0'].at['JND', 'PV_GW']:.1f}, "
    f"상한 {LIM['PV']['JND']:.1f}) 등이다. 상한 조건에서는 이 설비가 서울경기 PV({R['B0'].at['GGD', 'PV_GW']:.1f} GW)와 "
    '충북·경남·강원 PV, 그리고 여러 지역 풍력으로 흩어진다.'),
   ('설비 증가와 비용 증가: ', f"PV 이용률이 {U('B0', 'PV_CF_%'):.1f}% → {V('B0', 'PV_CF_%'):.1f}%로 조금 낮아지고 출력제한률이 "
    f"{U('B0', 'curtail_%'):.1f}% → {V('B0', 'curtail_%'):.1f}%로 올라, PV·풍력이 {RE_B0 - URE_B0:.1f} GW 더 필요하다. "
    f"원거리 설비를 수도권으로 보내야 해서 HVDC 수송량이 {U('B0', 'HVDC_TWh'):.0f} → {V('B0', 'HVDC_TWh'):.0f} TWh로 는다.")])

# 2장 결과
doc.add_page_break()
H('2. 비용 허용률별 최소 재생설비', 1)
F('fig2_capacity_mix.png', '그림 2. 잠재량 20% 상한 조건의 케이스별 PV·풍력 설비와 출력제한률 (검은 선: 상한 해제 총설비)')
T(compare_table(), font=7, caption='표 2-1. 같은 비용 허용률에서 상한 해제와 비교 (비용 증가율은 각 조건 B0 대비)')
T(mc_tbl, font=8, caption='표 2-2. 구간별 재생설비 감축 효율')
B([('0~1%: ', f"상한 해제와 비슷하게 효율적이다. 1 GW 감축당 비용은 {mc_cap[0]:,.0f} → {mc_cap[2]:,.0f}억원/년으로, 상한 해제"
    f"({mc_unc[0]:,.0f} → {mc_unc[2]:,.0f})와 큰 차이가 없다. 풍력이 대구경북에서 아직 늘 수 있고, "
    '상한 때문에 들어간 이용률 낮은 PV를 빼는 것이 싸기 때문이다.'),
   ('1~3%: ', f"1→2% 구간에서 {red_cap[3]:.1f} GW, 2→3% 구간에서 {red_cap[4]:.1f} GW만 줄어든다. 대구경북 풍력이 MGA-2에서 상한에 닿은 뒤로는 "
    '설비를 줄이는 수단이 열부문 전환·수전해 확대처럼 비싼 운영 조정만 남는다.'),
   ('출력제한: ', f"{V('B0', 'curtail_%'):.1f}% → MGA-1 {V(m1, 'curtail_%'):.1f}% → MGA-3 {V(m3, 'curtail_%'):.1f}%. "
    'PV 출력제한이 대부분이며, 줄어드는 폭은 상한 해제와 비슷하다.')])

# 3장 메커니즘
doc.add_page_break()
H('3. 감축 메커니즘', 1)
F('fig4_mechanism.png', '그림 3. 각 조건 B0 대비 연간 전력 수급 변화 (채색: 잠재량 20% 상한, 빗금: 상한 해제)')
B([('PV→풍력 전환의 한계: ', f"상한 조건 MGA-3의 풍력 증가는 {g(m3, 'WT_GW'):+.1f} GW(발전량 {g(m3, 'WT_TWh'):+.1f} TWh)로 "
    f"상한 해제({gu(m3, 'WT_GW'):+.1f} GW, {gu(m3, 'WT_TWh'):+.1f} TWh)보다 작다. 전국 풍력 상한 {WT_LIM:.0f} GW 중 "
    f"B0에서 이미 {V('B0', 'WT_GW'):.1f} GW를 쓰고 있어 늘릴 여지가 거의 없다."),
   ('열부문 전환이 일찍 등장: ', f"CO2 상한 때문에 가스 연료 총량({V('B0', 'gas_fuel_TWh'):.0f} TWh)은 고정된다. 같은 가스를 LNG 발전 대신 "
    f"CHP에 쓰면 CHP 열이 히트펌프 열을 대체해 전력 수요가 준다. 상한 조건에서는 MGA-1에서 이미 CHP 연료가 "
    f"{g(m1, 'CHP_fuel_TWh'):+.1f} TWh 늘고 히트펌프 전력이 {g(m1, 'HP_elec_TWh'):+.1f} TWh 줄어든다(상한 해제 MGA-1은 "
    f"{gu(m1, 'CHP_fuel_TWh'):+.1f}, {gu(m1, 'HP_elec_TWh'):+.1f} TWh)."),
   ('잉여 흡수 설비 확대 (2% 이상): ', f"수전해 설비가 {V('B0', 'electrolyser_GW'):.1f} → {V(m3, 'electrolyser_GW'):.1f} GW로 늘고 "
    f"MGA-3에서 수소 저장({V('B0', 'H2_storage_GWh'):.0f} → {V(m3, 'H2_storage_GWh'):.0f} GWh)과 열저장 투자도 늘어난다. "
    '남는 PV를 버리지 않고 쓰는 방향으로 PV 설비를 줄이는 것이다. 신규 ESS는 여전히 0이다.'),
   ('송전 이용 증가: ', f"상한 해제 MGA에서는 원거리 PV가 줄어 HVDC 수송량이 감소했지만, 상한 조건에서는 "
    f"{V('B0', 'HVDC_TWh'):.0f} → {V(m3, 'HVDC_TWh'):.0f} TWh로 늘어난다. 대구경북 풍력과 남부 PV를 수도권으로 보내야 하기 때문이다."),
   ('가격 지표 주의: ', f"CO2 잠재가격이 {V('B0', 'CO2_price'):.0f} → {V(m3, 'CO2_price'):,.0f}천원/t로 크게 오른다"
    f"(상한 해제 MGA-3 {U(m3, 'CO2_price'):.0f}). 재생 용량을 인위적으로 묶은 데 따른 잠재가격이므로 시장 가격 전망으로 해석하면 안 된다.")])
T(balance_table(), font=7.5, caption='표 3-1. 전력·열·가스 수급 변화 (TWh, 잠재량 20% 상한)', fmt='{:,.1f}')
F('fig5_cost_breakdown.png', '그림 4. 잠재량 20% 상한 조건의 B0 대비 비용 구성 변화')

# 4장 지역
doc.add_page_break()
H('4. 지역별 설비와 상한', 1)
F('fig3_regional.png', '그림 5. 지역별 PV(a)·풍력(b) 설비와 잠재량 20% 상한 (빨간 막대: 상한, ×: 상한 해제 B0)', width=17)
B([('상한에 걸린 지역: ', f"B0에서 PV는 {', '.join(REGION_KR[r] for r in binding('B0', 'PV'))}, 풍력은 "
    f"{', '.join(REGION_KR[r] for r in binding('B0', 'WT'))}이 상한에 걸려 있다. MGA-3에서도 PV 상한 지역은 그대로 상한을 유지한다."),
   ('PV 감축 지역: ', f"{pv_cut_text} GW (B0 → MGA-3). "
    '상한 해제 MGA에서 광주전남·대전세종충남·부산울산 PV가 줄던 것과 다르다.'),
   ('풍력 증가 지역: ', f"{', '.join(f'{REGION_KR[r]} {wt_add[r]:+.1f}' for r in wt_add_main)} GW (MGA-3). "
    '상한 해제에서 서울경기 풍력이 잠재량의 100%를 넘게 늘던 것과 달리, 모든 지역이 잠재량 20% 이내에 있다.')])
T(regional_table(), font=6.5, caption='표 4-1. 지역별 PV·풍력 설비와 상한 (GW)', fmt='{:,.1f}')

# 5장 시사점
doc.add_page_break()
H('5. 종합 및 시사점', 1)
B([('입지 제약이 있어도 1%까지는 유효하다: ', f"잠재량 20% 상한에서도 시스템 비용 +0.5%(약 {V(m05, 'cost_sys') - C_B0:.2f}조원/년)로 "
    f"{-dre(m05):.0f} GW, +1%(약 {V(m1, 'cost_sys') - C_B0:.2f}조원/년)로 {-dre(m1):.0f} GW를 줄일 수 있다. "
    '상한 해제 결과의 0~1% 구간 결론은 입지 제약을 걸어도 유지된다.'),
   ('그 이상은 풍력 입지가 관건이다: ', f"2% 이상에서 감축량이 상한 해제보다 크게 작아진다(3%에서 {-dre(m3):.1f} vs {-dre(m3, unc):.1f} GW). "
    '상한 해제 분석의 2~5% 감축량은 서울경기 풍력의 잠재량 초과를 전제로 한 값이므로, 정책 수치로는 상한 조건 결과를 쓰는 것이 타당하다.'),
   ('입지 제약 자체의 비용: ', f"잠재량 20% 상한은 B0 비용을 {C_B0 - U_B0:.2f}조원/년 올리고 필요한 PV·풍력을 {RE_B0 - URE_B0:.1f} GW 늘린다. "
    '좋은 입지(수도권 인근 풍력, 일사 좋은 PV 부지)의 확보가 보급 부담과 비용을 함께 줄이는 수단이다.'),
   ('수단 구성의 변화: ', '상한 해제에서는 PV→풍력 전환이 주 수단이었지만, 상한 조건에서는 풍력 전환 여지가 작아 CHP·히트펌프 등 '
    '열부문 운영과 수전해·열저장 같은 부문결합 설비가 더 일찍, 더 큰 비중으로 쓰인다. ESS 확대는 여전히 선택되지 않는다.'),
   ('한계: ', '① 잠재량 20%는 일률적 가정이며 지역별 실제 입지 가능량과 다를 수 있다. ② PV와 풍력 GW를 같은 가중치로 최소화하였다. '
    '③ MGA 대안의 가격 지표는 용량 제한에 따른 잠재가격이다. ④ 5% 이상 구간은 계산하지 않았다.')])

# 부록
doc.add_page_break()
H('부록. 산출 파일', 1)
B(['결과 폴더: results/mga_sweep_cap20_* (B0.nc, MGA-*.nc, mga_summary.csv, B0_reference.json)',
   f'분석 표: {OUT}/MGA_잠재량상한_분석_표.xlsx (요약·단계별·상한 해제 비교·구간 효율·B0 비교·수급·지역 + 전체 지표)',
   f'비교 대상: {args.base_dir} (상한 해제 MGA)',
   '실행 스크립트: run_mga_sweep.py --potential-cap 0.2 (MGA 실행), analyze_mga.py --out (집계), make_mga_cap_report.py (보고서)'])

path = os.path.join(OUT, 'MGA_잠재량20상한_분석_보고서.docx')
doc.save(path)
print(f'완료: {path}')
pd.set_option('display.width', 250)
print(main_table().round(3).to_string(index=False))
print(compare_table().round(2).to_string(index=False))
print(mc_tbl.round(1).to_string(index=False))
print(b0_table().round(2).to_string(index=False))
print('WT binding B0:', binding('B0', 'WT'), 'PV binding B0:', binding('B0', 'PV'))
print('PV binding MGA-3:', binding(m3, 'PV'), 'WT headroom MGA-3:', round(headroom(m3, 'WT'), 2))
