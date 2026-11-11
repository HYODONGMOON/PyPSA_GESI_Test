"""
민감도 분석 세트별 보고서(Word) + 그림(PNG) + 표(Excel) 생성

사용법:
  python make_sensitivity_report.py
출력:
  reports/sensitivity_<날짜>/민감도분석_보고서.docx, 민감도분석_표.xlsx, fig*.png
"""
import glob
import logging
import os
from datetime import datetime

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pypsa
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from run_regional_cost_sweep import TECH_POTENTIAL_GW, regional_multipliers

logging.disable(logging.WARNING)
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 9

R = 'results'
BASE_NC = rf"{R}\ESS_cost_sweep_20261028_095343\base_1.000\optimization_result_20261028_095935.nc"
OUT = os.path.join('reports', f"sensitivity_{datetime.now().strftime('%Y%m%d')}")
os.makedirs(OUT, exist_ok=True)

REGION_KR = {'BSN': '부산울산', 'CBD': '충북', 'CND': '대전세종충남', 'GBD': '대구경북', 'GGD': '서울경기',
             'GND': '경남', 'GWD': '강원', 'ICN': '인천', 'JBD': '전북', 'JJD': '제주', 'JND': '광주전남'}
C_PV, C_WT, C_ESS, C_LINE = '#F2B134', '#3C8DBC', '#7B5EA7', '#C0392B'


def kr_corridor(c):
    a, b = c.split('-')
    return f"{REGION_KR.get(a, a)}–{REGION_KR.get(b, b)}"


# ───────────────────────────── 데이터 ─────────────────────────────
def load_concat(pattern, fname, key='case'):
    dfs = []
    for d in sorted(glob.glob(pattern)):
        f = os.path.join(d, fname)
        if os.path.exists(f) and os.path.getsize(f) > 5:
            dfs.append(pd.read_csv(f))
    df = pd.concat(dfs, ignore_index=True)
    if 'PV_GW' in df.columns:
        df = df.dropna(subset=['PV_GW'])
    subset = [key] if fname == 'sweep_summary.csv' else [c for c in ['case', 'region', 'tech', 'corridor'] if c in df.columns]
    return df.drop_duplicates(subset, keep='last')


nb = pypsa.Network(BASE_NC)
g = nb.generators
base_reg = []
for tech in ('PV', 'WT'):
    for r, pot in TECH_POTENTIAL_GW[tech].items():
        nm = f'{r}_{tech}'
        base_reg.append({'case': 'base', 'region': r, 'tech': tech, 'existing_GW': g.at[nm, 'p_nom'] / 1e3,
                         'opt_GW': g.at[nm, 'p_nom_opt'] / 1e3, 'potential_GW': pot,
                         'use_%': g.at[nm, 'p_nom_opt'] / 1e3 / pot * 100})
base_reg = pd.DataFrame(base_reg)

ess = pd.read_csv(rf"{R}\ESS_cost_sweep_20261028_095343\sweep_summary.csv")
ess['RE_GW'] = ess.PV_GW + ess.WT_GW
ess['cost'] = ess.objective / 1e9
base = ess.iloc[0]

flex = pd.read_csv(rf"{R}\flex_sweep_20261028_113555\sweep_summary_true_cost.csv", index_col=0)
cde = pd.read_csv(rf"{R}\flex_sweep_CDE_20261028_142657\sweep_summary_extended.csv", index_col=0)
for c in cde.columns:
    if c in flex.columns:
        cde.loc['base', c] = flex.loc['base', c]

regp = pd.read_csv(rf"{R}\regional_cost_sweep_20261028_154525\sweep_summary.csv")
regp_cap = pd.read_csv(rf"{R}\regional_cost_sweep_20261028_154525\regional_capacity.csv")
mix = load_concat(rf"{R}\regional_cost_sweep_mix_*", 'sweep_summary.csv')
mix_cap = load_concat(rf"{R}\regional_cost_sweep_mix_*", 'regional_capacity.csv')
tr = load_concat(rf"{R}\transmission_sweep_*", 'sweep_summary.csv')
tr_corr = load_concat(rf"{R}\transmission_sweep_*", 'corridor_expansion.csv')
for df in (regp, mix, tr):
    df['RE_GW'] = df.PV_GW + df.WT_GW
    df['cost'] = df.objective / 1e9
tr['cost'] = tr.objective_incl_grid / 1e9


def newbuild_mult(caps, mult):
    out = {}
    for tech in ('PV', 'WT'):
        y = caps[caps.tech == tech].set_index('region')
        add = y.opt_GW - y.existing_GW
        out[tech] = float((add * mult[tech]).sum() / add.sum()) if add.sum() > 0 else np.nan
    return out


excel_tables = {}

# ───────────────────────────── 그림 ─────────────────────────────
figs = {}


def save(fig, name):
    p = os.path.join(OUT, name)
    fig.tight_layout()
    fig.savefig(p, dpi=200)
    plt.close(fig)
    figs[name] = p
    return p


# 세트1: ESS 비용
t1 = pd.DataFrame({
    'ESS 비용 배율': ['100%', '38.2% (손익분기)', '34.4% (−10%)', '30.5% (−20%)', '26.7% (−30%)'],
    'PV (GW)': ess.PV_GW, '풍력 (GW)': ess.WT_GW, '재생합계 (GW)': ess.RE_GW,
    'ESS 에너지 (GWh)': ess.ESS_GWh, 'ESS 출력 (GW)': ess.ESS_GW, 'ESS 충전량 (TWh)': ess.ESS_charge_TWh,
    '출력제한 (TWh)': ess.curtail_TWh, '출력제한률 (%)': ess['curtail_%'], 'V2G (TWh)': ess.V2G_TWh,
    'CO2 그림자가격 (천원/t)': ess.CO2_shadow_price, '총비용 (조원)': ess.cost})
excel_tables['1_ESS비용'] = t1
fig, ax = plt.subplots(1, 2, figsize=(9, 3.4))
x = np.arange(len(ess))
lbl = ['100%', '38.2%', '34.4%', '30.5%', '26.7%']
ax[0].bar(x, ess.PV_GW, color=C_PV, label='PV')
ax[0].bar(x, ess.WT_GW, bottom=ess.PV_GW, color=C_WT, label='풍력')
for i, v in enumerate(ess.RE_GW):
    ax[0].text(i, v + 5, f'{v:.0f}', ha='center', fontsize=8)
ax[0].set_xticks(x, lbl); ax[0].set_xlabel('ESS 비용 배율'); ax[0].set_ylabel('설비용량 (GW)')
ax[0].set_title('(a) 재생에너지 설비용량'); ax[0].legend(loc='lower left'); ax[0].set_ylim(0, 450)
ax[1].bar(x, ess.ESS_GWh, color=C_ESS, label='ESS 에너지 (GWh)')
ax[1].set_xticks(x, lbl); ax[1].set_xlabel('ESS 비용 배율'); ax[1].set_ylabel('ESS (GWh)')
a2 = ax[1].twinx(); a2.plot(x, ess['curtail_%'], 'o-', color=C_LINE, label='출력제한률')
a2.set_ylabel('출력제한률 (%)'); a2.set_ylim(0, 15)
ax[1].set_title('(b) ESS 보급량과 출력제한률')
h1, l1 = ax[1].get_legend_handles_labels(); h2, l2 = a2.get_legend_handles_labels()
ax[1].legend(h1 + h2, l1 + l2, loc='upper left', fontsize=8)
save(fig, 'fig1_ess_cost.png')

# 세트2: 출력제한 페널티
fa = flex.loc[['base', 'A_penalty_010', 'A_penalty_030', 'A_penalty_050', 'A_penalty_100']]
t2 = pd.DataFrame({
    '페널티 (천원/MWh)': [0, 10, 30, 50, 100], 'PV (GW)': fa.PV_GW.values, '풍력 (GW)': fa.WT_GW.values,
    '재생합계 (GW)': fa.RE_GW.values, '출력제한률 (%)': fa['curtail_%'].values,
    'ESS 충전량 (TWh)': fa.ESS_charge_TWh.values, '마이너스 가격 시간 (h)': fa.neg_price_hours.values,
    'CO2 그림자가격 (천원/t)': fa.CO2_shadow_price.values, '실제 총비용 (조원)': fa['true_cost_조'].values})
excel_tables['2_출력제한페널티'] = t2
fig, ax = plt.subplots(1, 2, figsize=(9, 3.4))
x = np.arange(5); lbl = ['0', '10', '30', '50', '100']
ax[0].bar(x, t2['PV (GW)'], color=C_PV, label='PV')
ax[0].bar(x, t2['풍력 (GW)'], bottom=t2['PV (GW)'], color=C_WT, label='풍력')
for i, v in enumerate(t2['재생합계 (GW)']):
    ax[0].text(i, v + 5, f'{v:.0f}', ha='center', fontsize=8)
a2 = ax[0].twinx(); a2.plot(x, t2['출력제한률 (%)'], 'o-', color=C_LINE, label='출력제한률'); a2.set_ylim(0, 14)
a2.set_ylabel('출력제한률 (%)')
ax[0].set_xticks(x, lbl); ax[0].set_xlabel('출력제한 페널티 (천원/MWh)'); ax[0].set_ylabel('설비용량 (GW)')
ax[0].set_ylim(0, 450); ax[0].set_title('(a) 설비용량과 출력제한률')
h1, l1 = ax[0].get_legend_handles_labels(); h2, l2 = a2.get_legend_handles_labels()
ax[0].legend(h1 + h2, l1 + l2, loc='lower left', fontsize=8)
dcost = t2['실제 총비용 (조원)'] - t2['실제 총비용 (조원)'][0]
dre = t2['재생합계 (GW)'][0] - t2['재생합계 (GW)']
ax[1].plot(dre, dcost, 'o-', color='k')
for i, p in enumerate(lbl):
    ax[1].annotate(f'c={p}', (dre[i], dcost[i]), textcoords='offset points', xytext=(5, -10), fontsize=8)
ax[1].set_xlabel('재생에너지 용량 감소 (GW)'); ax[1].set_ylabel('총비용 증가 (조원/년)')
ax[1].set_title('(b) 용량 감축 대비 비용 증가'); ax[1].grid(alpha=0.3)
save(fig, 'fig2_curtail_penalty.png')

# 세트3: 풍력 하한 + ESS 비용
fb = flex.loc[['base', 'B_windfloor_ess0.382', 'B_windfloor_ess0.305', 'B_windfloor_ess0.267']]
eo = ess.iloc[[0, 1, 3, 4]]
t3 = pd.DataFrame({
    'ESS 비용 배율': ['100%', '38.2%', '30.5%', '26.7%'],
    '[풍력고정] PV (GW)': fb.PV_GW.values, '[풍력고정] 풍력 (GW)': fb.WT_GW.values,
    '[풍력고정] 재생합계 (GW)': fb.RE_GW.values, '[풍력고정] 출력제한률 (%)': fb['curtail_%'].values,
    '[풍력고정] ESS (GWh)': fb.ESS_GWh.values, '[풍력고정] 총비용 (조원)': fb['true_cost_조'].values,
    '[ESS만] 재생합계 (GW)': eo.RE_GW.values, '[ESS만] 풍력 (GW)': eo.WT_GW.values,
    '[ESS만] 총비용 (조원)': eo.cost.values})
excel_tables['3_풍력하한+ESS'] = t3
fig, ax = plt.subplots(1, 2, figsize=(9, 3.2))
x = np.arange(4); lbl = ['100%', '38.2%', '30.5%', '26.7%']
ax[0].plot(x, t3['[ESS만] 재생합계 (GW)'], 's--', color='gray', label='ESS 비용만 인하')
ax[0].plot(x, t3['[풍력고정] 재생합계 (GW)'], 'o-', color=C_WT, label='풍력 80 GW 고정 + ESS 비용 인하')
ax[0].set_xticks(x, lbl); ax[0].set_xlabel('ESS 비용 배율'); ax[0].set_ylabel('재생에너지 합계 (GW)')
ax[0].set_title('(a) 재생에너지 설비용량'); ax[0].legend(fontsize=8); ax[0].grid(alpha=0.3)
ax[1].plot(x, t3['[ESS만] 총비용 (조원)'], 's--', color='gray', label='ESS 비용만 인하')
ax[1].plot(x, t3['[풍력고정] 총비용 (조원)'], 'o-', color=C_WT, label='풍력 고정 + ESS 비용 인하')
ax[1].set_xticks(x, lbl); ax[1].set_xlabel('ESS 비용 배율'); ax[1].set_ylabel('총비용 (조원/년)')
ax[1].set_title('(b) 총비용'); ax[1].legend(fontsize=8); ax[1].grid(alpha=0.3)
save(fig, 'fig3_windfloor_ess.png')

# 세트4: 조합 / 원자력
rows4 = [('기본', flex.loc['base']), ('페널티 30', flex.loc['A_penalty_030']),
         ('페널티 30 + ESS 38.2%', cde.loc['C_penalty030_ess0.382']),
         ('페널티 30 + ESS 30.5%', cde.loc['C_penalty030_ess0.305']),
         ('페널티 30 + ESS 26.7%', cde.loc['C_penalty030_ess0.267']),
         ('원자력 최소출력 0.6', cde.loc['D_nucpmin_0.6']), ('원자력 최소출력 0.4', cde.loc['D_nucpmin_0.4']),
         ('원자력 0.6 + 페널티 30', cde.loc['E_nucpmin_0.6_penalty030'])]
t4 = pd.DataFrame([{'케이스': k, 'PV (GW)': v.PV_GW, '풍력 (GW)': v.WT_GW, '재생합계 (GW)': v.RE_GW,
                    '출력제한률 (%)': v['curtail_%'], 'ESS (GWh)': v.ESS_GWh, '원자력 발전량 (TWh)': v.nuc_TWh,
                    '재생 발전량 (TWh)': v.RE_gen_TWh, 'CO2 그림자가격 (천원/t)': v.CO2_shadow_price,
                    '실제 총비용 (조원)': v['true_cost_조']} for k, v in rows4])
excel_tables['4_조합_원자력'] = t4
fig, ax = plt.subplots(1, 2, figsize=(9.5, 3.6))
y = np.arange(len(t4))[::-1]
ax[0].barh(y, t4['PV (GW)'], color=C_PV, label='PV')
ax[0].barh(y, t4['풍력 (GW)'], left=t4['PV (GW)'], color=C_WT, label='풍력')
for yi, v in zip(y, t4['재생합계 (GW)']):
    ax[0].text(v + 3, yi, f'{v:.0f}', va='center', fontsize=8)
ax[0].set_yticks(y, t4['케이스']); ax[0].set_xlim(0, 450); ax[0].set_xlabel('설비용량 (GW)')
ax[0].set_title('(a) 재생에너지 설비용량')
ax[0].legend(fontsize=8, loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=2)
ax[1].barh(y, t4['출력제한률 (%)'], color=C_LINE, alpha=0.8)
for yi, v in zip(y, t4['출력제한률 (%)']):
    ax[1].text(v + 0.2, yi, f'{v:.1f}%', va='center', fontsize=8)
ax[1].set_yticks(y, [''] * len(y)); ax[1].set_xlim(0, 14); ax[1].set_xlabel('출력제한률 (%)')
ax[1].set_title('(b) 출력제한률')
save(fig, 'fig4_combo_nuclear.png')

# 세트5: 지역 비용 차등 (절대 잠재량)
cap5 = pd.concat([base_reg, regp_cap])
p5 = cap5.pivot_table(index=['tech', 'region'], columns='case', values='opt_GW')[
    ['base', 'regcost_alpha0.25', 'regcost_alpha0.50', 'regcost_alpha0.75']]
p5.columns = ['기본', 'α=0.25', 'α=0.5', 'α=0.75']
mult5 = {a: regional_multipliers(nb, a, 'potential') for a in (0.25, 0.5, 0.75)}
t5a = p5.round(1).reset_index()
t5a['region'] = t5a.region.map(REGION_KR)
t5a.insert(2, '기술적 잠재량 (GW)', [TECH_POTENTIAL_GW.at[r, t] for t, r in p5.index])
t5a = t5a.rename(columns={'tech': '구분', 'region': '지역'})
nb5 = {f'regcost_alpha{a:.2f}': newbuild_mult(regp_cap[regp_cap.case == f'regcost_alpha{a:.2f}'], mult5[a])
       for a in (0.25, 0.5, 0.75)}
t5b = pd.DataFrame({
    '케이스': ['기본', 'α=0.25', 'α=0.5', 'α=0.75'],
    'PV (GW)': [base.PV_GW] + list(regp.PV_GW), '풍력 (GW)': [base.WT_GW] + list(regp.WT_GW),
    '재생합계 (GW)': [base.RE_GW] + list(regp.RE_GW), '출력제한률 (%)': [base['curtail_%']] + list(regp['curtail_%']),
    '총비용 (조원)': [base.cost] + list(regp.cost),
    '신규 PV 평균 배율': [1.0] + [nb5[c]['PV'] for c in regp.case],
    '신규 풍력 평균 배율': [1.0] + [nb5[c]['WT'] for c in regp.case]})
t5m = pd.concat({f'α={a}': mult5[a] for a in mult5}, axis=1).round(2)
t5m.columns = [f"{a} {'PV' if t == 'PV' else '풍력'}" for a, t in t5m.columns]
t5m.index = t5m.index.map(REGION_KR)
excel_tables['5_지역비용차등_전국'] = t5b
excel_tables['5_지역비용차등_지역별'] = t5a
excel_tables['5_지역비용차등_배율'] = t5m.reset_index().rename(columns={'index': '지역'})
fig, ax = plt.subplots(2, 1, figsize=(9, 5.2))
cols = ['#888888', '#9ECAE1', '#4292C6', '#08519C']
for k, tech in enumerate(('PV', 'WT')):
    d = p5.loc[tech]
    xx = np.arange(len(d)); w = 0.2
    for j, c in enumerate(d.columns):
        ax[k].bar(xx + (j - 1.5) * w, d[c], w, color=cols[j], label=c)
    ax[k].plot(xx, [TECH_POTENTIAL_GW.at[r, tech] * 0.2 for r in d.index], '_', color='r', ms=14, mew=2,
               label='잠재량의 20%')
    ax[k].set_xticks(xx, [REGION_KR[r] for r in d.index], fontsize=8)
    ax[k].set_ylabel('GW'); ax[k].set_title(f"({'a' if k == 0 else 'b'}) 지역별 {'PV' if tech == 'PV' else '풍력'} 설비용량")
    ax[k].legend(fontsize=7, ncol=5)
save(fig, 'fig5_regional_cost.png')

# 세트6: 상한·1인당 잠재량·수도권
order6 = [('base', '기본'), ('cap20_uniform', '상한20%'), ('cap20_regcost_alpha0.25', '상한20%+절대α0.25'),
          ('percap_alpha0.25', '1인당α0.25'), ('percap_alpha0.50', '1인당α0.5'), ('metro_w1.3', '수도권×1.3'),
          ('cap20_percap_alpha0.25', '상한20%+1인당α0.25')]
cap6 = pd.concat([base_reg, mix_cap])
use6 = cap6.pivot_table(index=['tech', 'region'], columns='case', values='use_%')[[c for c, _ in order6]]
use6.columns = [k for _, k in order6]
opt6 = cap6.pivot_table(index=['tech', 'region'], columns='case', values='opt_GW')[[c for c, _ in order6]]
opt6.columns = [k for _, k in order6]
mx = mix.set_index('case')
t6 = pd.DataFrame([{
    '케이스': k, 'PV (GW)': base.PV_GW if c == 'base' else mx.at[c, 'PV_GW'],
    '풍력 (GW)': base.WT_GW if c == 'base' else mx.at[c, 'WT_GW'],
    '재생합계 (GW)': base.RE_GW if c == 'base' else mx.at[c, 'RE_GW'],
    '출력제한률 (%)': base['curtail_%'] if c == 'base' else mx.at[c, 'curtail_%'],
    '총비용 (조원)': base.cost if c == 'base' else mx.at[c, 'cost'],
    'CO2 그림자가격 (천원/t)': base.CO2_shadow_price if c == 'base' else mx.at[c, 'CO2_shadow_price'],
    '잠재량 대비 최대 지역 (%)': use6[k].max(),
    '최대 지역': f"{REGION_KR[use6[k].idxmax()[1]]} {'PV' if use6[k].idxmax()[0] == 'PV' else '풍력'}"}
    for c, k in order6])
excel_tables['6_상한_1인당_수도권_전국'] = t6
t6r = opt6.round(1).reset_index()
t6r['region'] = t6r.region.map(REGION_KR)
t6r = t6r.rename(columns={'tech': '구분', 'region': '지역'})
excel_tables['6_상한_1인당_수도권_지역별'] = t6r
u6 = use6.round(0).reset_index(); u6['region'] = u6.region.map(REGION_KR)
excel_tables['6_잠재량대비이용률(%)'] = u6.rename(columns={'tech': '구분', 'region': '지역'})
fig, ax = plt.subplots(1, 2, figsize=(10.5, 4.2), gridspec_kw={'wspace': 0.35})
for k, tech in enumerate(('PV', 'WT')):
    d = use6.loc[tech]
    im = ax[k].imshow(d.values, cmap='YlOrRd', vmin=0, vmax=100, aspect='auto')
    ax[k].set_yticks(range(len(d)), [REGION_KR[r] for r in d.index], fontsize=8)
    ax[k].set_xticks(range(d.shape[1]), d.columns, rotation=40, ha='right', fontsize=7)
    for i in range(d.shape[0]):
        for j in range(d.shape[1]):
            v = d.values[i, j]
            ax[k].text(j, i, f'{v:.0f}', ha='center', va='center', fontsize=6.5,
                       color='white' if v > 60 else 'black')
    ax[k].set_title(f"({'a' if k == 0 else 'b'}) {'PV' if tech == 'PV' else '풍력'} 설비 / 기술적 잠재량 (%)")
fig.colorbar(im, ax=ax, shrink=0.8, label='%')
fig.savefig(os.path.join(OUT, 'fig6_cap_percap_metro.png'), dpi=200, bbox_inches='tight')
plt.close(fig)
figs['fig6_cap_percap_metro.png'] = os.path.join(OUT, 'fig6_cap_percap_metro.png')

# 세트7: 송전망
order7 = [('ref_nogrid', '확장 없음'), ('uniform_10', '일괄 +10%'), ('uniform_20', '일괄 +20%'),
          ('uniform_30', '일괄 +30%'), ('need_cost0.5', '필요구간 (건설비×0.5)'),
          ('need_cost1.0', '필요구간 (건설비×1)'), ('need_cost2.0', '필요구간 (건설비×2)')]
tt = tr.set_index('case')
t7 = pd.DataFrame([{'케이스': k, '송전 증설 (GW)': tt.at[c, 'grid_added_GW'],
                    '송전 연간비용 (조원)': tt.at[c, 'grid_annual_cost_bn'] / 1e3,
                    'PV (GW)': tt.at[c, 'PV_GW'], '풍력 (GW)': tt.at[c, 'WT_GW'], '재생합계 (GW)': tt.at[c, 'RE_GW'],
                    '출력제한률 (%)': tt.at[c, 'curtail_%'], 'CO2 그림자가격 (천원/t)': tt.at[c, 'CO2_shadow_price'],
                    '가스 발전 (TWh)': tt.at[c, 'gas_gen_TWh'], '총비용·송전포함 (조원)': tt.at[c, 'cost']}
                   for c, k in order7])
excel_tables['7_송전망_전국'] = t7
pc = tr_corr.pivot_table(index=['type', 'corridor'], columns='case', values='added_MW')
pc = (pc[[c for c, _ in order7]] / 1e3).round(2)
pc.columns = [k for _, k in order7]
ex = tr_corr.drop_duplicates(['type', 'corridor']).set_index(['type', 'corridor'])['existing_MW'] / 1e3
pc.insert(0, '기존 (GW)', ex.round(1))
hc = tr_corr.pivot_table(index=['type', 'corridor'], columns='case', values='congested_h')[['ref_nogrid', 'need_cost1.0']]
pc['혼잡시간 확장없음 (h)'] = hc['ref_nogrid']
pc['혼잡시간 필요구간×1 (h)'] = hc['need_cost1.0']
t7c = pc.reset_index()
t7c['corridor'] = t7c.corridor.map(kr_corridor)
t7c = t7c.rename(columns={'type': '유형', 'corridor': '구간'})
excel_tables['7_송전망_구간별(GW)'] = t7c
fig, ax = plt.subplots(1, 2, figsize=(10, 3.8))
need = pc[['필요구간 (건설비×0.5)', '필요구간 (건설비×1)', '필요구간 (건설비×2)']]
need = need[need.max(axis=1) > 0.1].sort_values('필요구간 (건설비×1)')
yy = np.arange(len(need)); w = 0.27
for j, (c, col) in enumerate(zip(need.columns, ['#9ECAE1', '#3182BD', '#08519C'])):
    ax[0].barh(yy + (j - 1) * w, need[c], w, color=col, label=c.replace('필요구간 ', ''))
ax[0].set_yticks(yy, [kr_corridor(c) for _, c in need.index], fontsize=8)
ax[0].set_xlabel('증설 용량 (GW)'); ax[0].set_title('(a) 필요 구간 확장: 구간별 증설량'); ax[0].legend(fontsize=7)
for _, r in t7.iterrows():
    col = '#3182BD' if '필요' in r['케이스'] else ('#E6550D' if '일괄' in r['케이스'] else 'k')
    ax[1].scatter(r['송전 증설 (GW)'], r['총비용·송전포함 (조원)'], color=col, s=35)
    ax[1].annotate(r['케이스'].replace('필요구간 ', ''), (r['송전 증설 (GW)'], r['총비용·송전포함 (조원)']),
                   textcoords='offset points', xytext=(4, 4), fontsize=7)
ax[1].set_xlim(-3, 72)
ax[1].set_xlabel('송전 증설 (GW)'); ax[1].set_ylabel('총비용·송전 포함 (조원/년)')
ax[1].set_title('(b) 송전 증설량과 총비용'); ax[1].grid(alpha=0.3)
save(fig, 'fig7_transmission.png')

# 종합 그림: 전 케이스 재생합계 vs 출력제한률
pts = []
pts += [('ESS 비용', r.RE_GW, r['curtail_%']) for _, r in ess.iterrows()]
pts += [('출력제한 페널티', r['재생합계 (GW)'], r['출력제한률 (%)']) for _, r in t2.iloc[1:].iterrows()]
pts += [('풍력 고정+ESS', r['[풍력고정] 재생합계 (GW)'], r['[풍력고정] 출력제한률 (%)']) for _, r in t3.iloc[1:].iterrows()]
pts += [('조합·원자력', r['재생합계 (GW)'], r['출력제한률 (%)']) for _, r in t4.iloc[2:].iterrows()]
pts += [('지역 비용 차등', r['재생합계 (GW)'], r['출력제한률 (%)']) for _, r in t5b.iloc[1:].iterrows()]
pts += [('상한·1인당·수도권', r['재생합계 (GW)'], r['출력제한률 (%)']) for _, r in t6.iloc[1:].iterrows()]
pts += [('송전망', r['재생합계 (GW)'], r['출력제한률 (%)']) for _, r in t7.iterrows()]
pdf = pd.DataFrame(pts, columns=['세트', '재생합계 (GW)', '출력제한률 (%)'])
excel_tables['0_전케이스_산점도'] = pdf
fig, ax = plt.subplots(figsize=(7, 4.2))
for (s, d), mk in zip(pdf.groupby('세트', sort=False), ['o', 's', '^', 'D', 'v', 'P', 'X']):
    ax.scatter(d['재생합계 (GW)'], d['출력제한률 (%)'], marker=mk, s=40, label=s, alpha=0.85)
ax.scatter([base.RE_GW], [base['curtail_%']], marker='*', s=220, color='k', label='기본 케이스', zorder=5)
ax.set_xlabel('재생에너지 설비용량 (GW)'); ax.set_ylabel('출력제한률 (%)'); ax.grid(alpha=0.3)
ax.set_title('민감도 분석 전 케이스: 재생에너지 용량 vs 출력제한률')
ax.legend(fontsize=7, ncol=2)
save(fig, 'fig0_overview.png')

# ───────────────────────────── Word ─────────────────────────────
doc = Document()
st = doc.styles['Normal']
st.font.name = 'Malgun Gothic'; st.font.size = Pt(10)
st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Malgun Gothic')
for s in ('Heading 1', 'Heading 2', 'Heading 3', 'Title'):
    hs = doc.styles[s]
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
    if align:
        p.alignment = align
    p.paragraph_format.space_after = Pt(3)
    return p


def B(items, style='List Bullet'):
    for it in items:
        p = doc.add_paragraph(style=style)
        if isinstance(it, tuple):
            r = p.add_run(it[0]); r.bold = True
            p.add_run(it[1])
        else:
            p.add_run(it)
        p.paragraph_format.space_after = Pt(1)


def T(df, fmt=None, font=8, col_w=None, caption=None):
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
                f = (fmt or {}).get(c, '{:,.1f}')
                s = '' if pd.isna(v) else f.format(v)
            else:
                s = str(v)
            cells[j].text = ''
            rr = cells[j].paragraphs[0].add_run(s); rr.font.size = Pt(font)
            cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if j else WD_ALIGN_PARAGRAPH.LEFT
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def F(name, caption, width=16.5):
    doc.add_picture(figs[name], width=Cm(width))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    P(caption, italic=True, size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)


def H(text, lv):
    doc.add_heading(text, lv)


# 표지
doc.add_heading('PyPSA 전력계통 모형 민감도 분석 결과', 0)
P('재생에너지 과다보급·출력제한 대응 방안 및 지역 입지 현실화 검토', size=12)
P(f"작성일: {datetime.now():%Y-%m-%d}   |   분석 기간: 2026-10-28 ~ 2026-10-29", size=9)

H('개요', 1)
P('본 문서는 현재 입력조건(integrated_input_data.xlsx, 2026-10-28 재생성본)을 기준으로 수행한 7개 민감도 분석 세트의 '
  '결과를 정리한 것이다. 모든 케이스는 동일한 CO2 배출 상한(58.87 MtCO2)을 만족하도록 최적화되었으며, '
  'interface 및 통합 입력 파일은 수정하지 않고 메모리 상에서만 조건을 변경하였다.')
T(pd.DataFrame([
    ['기본 케이스', 'PV 299 GW, 풍력 80 GW(합계 379 GW), 출력제한 11.4%, ESS 107 GWh, 총비용 59.80조원/년'],
    ['CO2 제약', '배출 상한 58.87 MtCO2 (구속 조건, 그림자가격 329 천원/tCO2)'],
    ['최적화', 'PyPSA 0.31.1 + linopy, CPLEX Barrier(Crossover 생략), 8,784시간 단일연도'],
    ['CHP', '기존 CHP 배출 산정 및 제약 유지(변경 없음)'],
    ['비용 단위', '총비용 = 연간화 투자비 + 운전비(기존 설비 고정비는 상수로 제외), 천원 단위 입력 → 조원 환산'],
], columns=['항목', '내용']), font=8.5)

T(pd.DataFrame([
    ['1', 'ESS 비용', 'ESS 비용을 손익분기점(38.2%)에서 −10/−20/−30%', 'ESS는 비용을 낮추지만 PV를 늘려 재생 용량은 증가(379→405 GW)'],
    ['2', '출력제한 페널티', '버려지는 전력에 10~100 천원/MWh 부과', 'PV↓·풍력↑로 재생 용량 감소(379→349 GW), 비용 소폭 증가'],
    ['3', '풍력 하한 + ESS 비용', '풍력 80 GW 고정 + ESS 비용 인하', 'ESS가 PV 잉여만 흡수 → 재생 364 GW, 비용 감소'],
    ['4', '조합·원자력 유연성', '페널티+ESS, 원자력 최소출력 0.8→0.6/0.4', '원자력 유연성은 출력제한률만 낮추고 용량은 불변'],
    ['5', '지역 비용 차등(절대 잠재량)', '기술적 잠재량 기반 지역별 비용 배율 α=0.25~0.75', '인천·부산 과다보급 해소, 서울경기로 이동'],
    ['6', '상한·1인당 잠재량·수도권', '잠재량 20% 상한, 1인당 잠재량 배율, 수도권 ×1.3/1.5', '상한이 있어야 지역 집중 방지, 재생 용량은 증가(393 GW)'],
    ['7', '송전망 확장', '필요 구간 확장 vs 일괄 10/20/30%', '5개 구간 35 GW로 총비용 −1.5조, 용량 감소는 5 GW 수준'],
], columns=['No.', '세트', '변경 조건', '핵심 결과']), font=8, caption='표 0. 민감도 분석 세트 요약')
F('fig0_overview.png', '그림 0. 전 케이스의 재생에너지 설비용량과 출력제한률 분포')
H('핵심 메시지', 2)
B([('재생에너지 용량은 CO2 제약이 결정한다. ', '모든 케이스에서 필요한 무탄소 발전량(재생 ~500 TWh + 원자력)이 거의 같으므로, '
    '용량은 PV/풍력 구성비와 출력제한률로만 달라진다(PV 이용률 0.15 < 풍력 0.25).'),
   ('유연성 자원 단독으로는 용량이 줄지 않는다. ', 'ESS·V2G 비용 인하는 PV의 가치를 높여 PV가 풍력을 대체하고, '
    '원자력 유연성은 재생 출력제한 대신 원자력 감발로 대체될 뿐이다.'),
   ('PV/풍력 구성비를 바꾸는 신호가 필요하다. ', '출력제한 페널티 또는 풍력 하한과 결합할 때만 ESS가 용량 감소로 이어진다.'),
   ('지역 입지 현실화는 상한이 핵심이며 비용이 따른다. ', '상한 20% + 1인당 잠재량 차등 시 재생 용량은 393 GW로 늘고 '
    '출력제한도 12.8%로 증가한다.'),
   ('송전망은 "필요 구간" 확장이 효율적이다. ', '5개 구간 35 GW 증설(연 0.16조원)로 총비용이 1.5조원 감소하나, '
    '봄철 전국 동시 잉여에 따른 출력제한(12%대)은 송전으로 해결되지 않는다.')])

# 세트 1
doc.add_page_break()
H('세트 1. ESS 비용 민감도', 1)
H('분석 목적 및 설정', 2)
B(['ESS 비용을 얼마나 낮추면 출력제한과 재생에너지 과다보급이 줄어드는지 검토',
   '가격수용자(price-taker) 차익거래 LP로 ESS 지속시간별(1~8시간) 연간 가치를 추정 → 손익분기 비용 배율 k* = 0.382 '
   '(6시간 기준 가치 약 46,000 vs 비용 125,600 천원/MW·년)',
   'ESS Store 및 충·방전 인버터(Link) capital_cost에 배율 100% / 38.2% / 34.4% / 30.5% / 26.7% 적용'])
T(t1.drop(columns=['ESS 충전량 (TWh)', 'V2G (TWh)']), fmt={'총비용 (조원)': '{:.2f}', '출력제한률 (%)': '{:.1f}'},
  caption='표 1. ESS 비용 배율별 결과')
F('fig1_ess_cost.png', '그림 1. ESS 비용 배율별 재생에너지 용량(a)과 ESS 보급량·출력제한률(b)')
H('주요 결과 및 해석', 2)
B([('손익분기 추정 검증: ', '배율 38.2%에서 ESS가 추가되기 시작(107→133 GWh)하여 추정치와 일치'),
   ('재생 용량 증가: ', 'ESS가 359 GWh까지 늘어도 PV 299→368 GW, 풍력 80→37 GW로 재구성되어 합계는 379→405 GW로 증가'),
   ('출력제한 불변: ', '출력제한률은 11.1~12.0%로 거의 변화 없음. 재생 발전량(약 500 TWh)과 가스 발전(158.6 TWh)도 동일'),
   ('비용 감소: ', '총비용 59.80→57.87조원, CO2 그림자가격 329→217 천원/t로 CO2 감축 부담은 크게 완화'),
   ('해석: ', 'ESS는 낮 시간 PV 잉여를 저녁으로 옮겨 PV의 시스템 가치를 높인다. 이용률이 낮은 PV가 이용률이 높은 '
    '풍력을 대체하므로 같은 발전량을 위해 더 많은 GW가 필요해진다.')])
H('시사점 (시나리오 반영)', 2)
B(['ESS 비용 하락 시나리오는 "비용 절감·CO2 감축 부담 완화" 효과로 제시하고, "재생 과다보급 해소" 수단으로 제시하지 않는 것이 적절',
   'ESS 보급 확대를 용량 감소로 연결하려면 풍력 비중 유지(세트 3) 또는 출력제한 비용 신호(세트 2)와 함께 설계 필요',
   '손익분기 배율(현재 비용의 약 38%)은 ESS 보급 목표 시나리오의 비용 전제를 설정할 때 기준값으로 활용 가능'])

# 세트 2
doc.add_page_break()
H('세트 2. 출력제한 페널티 (마이너스 가격 신호)', 1)
H('분석 목적 및 설정', 2)
B(['출력제한 자체에 비용을 부과하여 실제 시장의 마이너스 가격처럼 잉여 시간대 가격 신호를 반영',
   '구현: PV·풍력에 capital_cost += c × Σ시간 가용률(연간 가용 시간), marginal_cost −= c → 목적함수에 c × 출력제한량이 더해지는 것과 동일',
   '페널티 c = 10 / 30 / 50 / 100 천원/MWh (ESS 비용은 현재 수준)',
   '참고 사례: PyPSA-Eur의 curtailment_mode, PLEXOS ISO-NE 출력제한 임계가격(−10 $/MWh), NREL 출력제한·PTC 분석'])
T(t2, fmt={'실제 총비용 (조원)': '{:.2f}', '마이너스 가격 시간 (h)': '{:,.0f}', '페널티 (천원/MWh)': '{:.0f}'},
  caption='표 2. 출력제한 페널티별 결과 (실제 총비용 = 페널티 금액 제외)')
F('fig2_curtail_penalty.png', '그림 2. 페널티별 설비용량·출력제한률(a)과 용량 감축 대비 비용 증가(b)')
H('주요 결과 및 해석', 2)
B([('용량 감소: ', '페널티 30에서 재생 363 GW(−16 GW), 출력제한 8.0%. 페널티 100에서 349 GW, 5.6%'),
   ('풍력 자연 증가: ', '하한 없이도 풍력 80→90 GW로 증가. 출력제한이 많은 PV가 불리해지고 이용률이 높은 풍력이 PV를 대체'),
   ('기존 ESS 활용 증가: ', '잉여 시간대 가격이 −c로 내려가 충·방전 가격차가 커지면서 ESS 충전량 33→53 TWh. 신규 ESS 투자는 발생하지 않음'),
   ('비용: ', '페널티 30에서 +0.26조원(0.4%)으로 16 GW 감축. 페널티 100은 +1.05조원으로 한계비용 급증'),
   ('주의: ', 'CO2 그림자가격 상승(329→449)은 페널티가 재생 비용처럼 작동한 결과로, 실제 감축비용 상승이 아님')])
H('시사점 (시나리오 반영)', 2)
B(['출력제한에 대한 가격 신호(마이너스 가격, 출력제한 보상 축소 등)는 재생 용량을 줄이는 직접적 수단',
   '페널티 30 천원/MWh 내외가 비용 대비 효율적인 구간 → "시장 가격신호 반영" 시나리오의 기준값으로 제안',
   '보고서에는 "출력제한률 5~8% 수준으로 관리 시 재생 용량 15~30 GW 절감 가능, 비용 증가 0.3~1.0조원/년"으로 제시 가능'])

# 세트 3
doc.add_page_break()
H('세트 3. 풍력 하한 유지 + ESS 비용 인하', 1)
H('분석 목적 및 설정', 2)
B(['ESS 비용이 낮아질 때 풍력이 밀려나는 현상을 막으면 출력제한과 재생 용량이 함께 줄어드는지 검토',
   '지역별 풍력 p_nom_min = 기본 케이스 풍력 용량(합계 80 GW) + ESS 비용 배율 100% / 38.2% / 30.5% / 26.7%',
   '비교 대상: 세트 1(ESS 비용만 인하)'])
T(t3, fmt={'[풍력고정] 총비용 (조원)': '{:.2f}', '[ESS만] 총비용 (조원)': '{:.2f}'}, font=7,
  caption='표 3. 풍력 고정 여부에 따른 결과 비교')
F('fig3_windfloor_ess.png', '그림 3. 풍력 고정 여부에 따른 재생에너지 용량(a)과 총비용(b)')
H('주요 결과 및 해석', 2)
B([('용량 감소: ', 'ESS 비용 26.7%에서 재생 합계 364 GW(ESS만 인하 시 405 GW 대비 −41 GW), 출력제한 9.3%'),
   ('ESS 역할 변화: ', '풍력을 고정하면 ESS(156 GWh)가 풍력을 대체하지 않고 PV 잉여만 흡수'),
   ('비용: ', '총비용 59.43조원으로 기본 대비 −0.37조원. 다만 같은 ESS 비용에서 풍력을 고정하지 않은 경우(57.87조원)보다 1.56조원 높음 '
    '→ 총 용량 감축의 대가'),
   ('해석: ', '풍력 하한은 PV/풍력 구성비를 고정하는 정책 제약으로, ESS가 "재배치"가 아닌 "출력제한 흡수"에 쓰이도록 만든다')])
H('시사점 (시나리오 반영)', 2)
B(['"풍력 보급목표 유지 + ESS 확대" 시나리오는 재생 용량 감축과 비용 절감을 동시에 달성하는 가장 균형적인 조합',
   '풍력 보급목표(해상풍력 로드맵 등)를 하한으로 두는 것은 정책 시나리오 설계 방식과도 부합',
   '풍력 하한 유지 비용(최대 1.6조원/년)은 "에너지 안보·계절 보완성" 등 모형 외 가치와 비교하여 제시 필요'])

# 세트 4
doc.add_page_break()
H('세트 4. 조합 및 원자력 유연성', 1)
H('분석 목적 및 설정', 2)
B(['C. 출력제한 페널티 30 + ESS 비용 인하(38.2% / 30.5% / 26.7%)',
   'D. 원자력 최소출력(p_min_pu) 0.8 → 0.6 / 0.4 (원자력 35.2 GW, 최대출력 0.85)',
   'E. 원자력 최소출력 0.6 + 페널티 30'])
T(t4, fmt={'실제 총비용 (조원)': '{:.2f}'}, font=7.5, caption='표 4. 조합·원자력 유연성 케이스 결과')
F('fig4_combo_nuclear.png', '그림 4. 케이스별 재생에너지 용량(a)과 출력제한률(b)')
H('주요 결과 및 해석', 2)
B([('페널티 + ESS 손익분기(38.2%): ', '재생 365 GW, 출력제한 7.9%, ESS 140 GWh로 두 효과가 더해짐'),
   ('ESS가 더 저렴해지면: ', '페널티 30으로는 PV의 풍력 대체를 막지 못해 풍력 37 GW, 재생 393 GW로 다시 증가'),
   ('원자력 유연성: ', '출력제한률 11.4→7.5%로 감소하지만 재생 용량은 379~380 GW로 불변. 원자력 발전량 259→235 TWh 감소분을 '
    '재생 발전(501→524 TWh)이 대체'),
   ('원자력 + 페널티: ', '재생 364 GW, 출력제한 6.4%(전체 최저), 총비용 59.95조원(페널티 단독 대비 −0.11조원)'),
   ('해석: ', '출력제한률이 낮아졌다고 과다보급이 해소된 것은 아니다. 원자력 감발은 무탄소 발전원 간의 대체일 뿐이다')])
H('시사점 (시나리오 반영)', 2)
B(['출력제한률은 보조 지표로 쓰고, 재생 용량·총비용과 함께 제시해야 오해를 피할 수 있음',
   '원자력 유연운전은 출력제한 완화·운영비 절감 수단으로 시나리오에 포함하되, 재생 보급 규모 결정과는 분리하여 해석',
   '"페널티 30 + ESS 손익분기 비용 + 원자력 0.6"을 유연성 종합 시나리오 후보로 검토 가능(미실행 조합)'])

# 세트 5
doc.add_page_break()
H('세트 5. 지역별 재생에너지 비용 차등 (절대 기술적 잠재량 기준)', 1)
H('분석 목적 및 설정', 2)
B(['인천·부산 등 면적이 작고 도시화된 지역에 PV가 과다 보급되는 문제 해소',
   '지역 배율 m = k × (지역 기술적 잠재량 / 지역 평균)^(−α), k는 기존 설비용량 가중평균 배율 = 1이 되도록 설정',
   '잠재량: PV = KIER 신재생에너지자원지도 기술적 잠재량(합계 1,807 GW), 풍력 = KIER 육상(352 GW) + 해양공간계획 유보해역 해상(123.5 GW)',
   'α = 0.25 / 0.5 / 0.75, PV·풍력 모두 적용'])
T(t5b, fmt={'총비용 (조원)': '{:.2f}', '신규 PV 평균 배율': '{:.2f}', '신규 풍력 평균 배율': '{:.2f}'},
  caption='표 5-1. 전국 결과 (신규 평균 배율 = 신규 설비용량 가중 비용 배율)')
F('fig5_regional_cost.png', '그림 5. 지역별 PV(a)·풍력(b) 설비용량 (빨간 표시: 기술적 잠재량의 20%)')
H('주요 결과 및 해석', 2)
B([('과다보급 해소: ', 'α=0.25만으로 인천 PV 51.6→1.2 GW(기존 설비만), 부산울산 29.5→2.9 GW. 풍력도 서울경기·충북에서 '
    '광주전남·대구경북·강원으로 이동'),
   ('새로운 집중: ', '경기의 기술적 잠재량(234 GW)이 커서 서울경기 배율이 1보다 낮아지고, 서울경기 PV가 16→100 GW(α=0.75, 잠재량의 40%)로 증가'),
   ('비용·용량 해석 주의: ', '신규 설비 가중 배율이 PV 0.84~0.96, 풍력 0.53~0.87로 평균 비용이 낮아진 효과가 섞임. '
    '순수 입지 효과는 α=0.25 수준(재생 379→378 GW)'),
   ('출력제한: ', '남부 집중으로 11.6~11.9%로 소폭 증가')])
H('시사점 (시나리오 반영)', 2)
B(['절대 기술적 잠재량은 도시화·수용성을 충분히 반영하지 못함 → 단독 지표로는 부적합',
   '비용 차등 시 정규화 기준(기존 설비 vs 신규 설비)에 따라 총비용이 크게 달라지므로 보고서에 기준을 명시해야 함'])
T(t5m.reset_index().rename(columns={'index': '지역'}), fmt={c: '{:.2f}' for c in t5m.columns}, font=7.5,
  caption='표 5-2. 지역별 비용 배율 (PV / 풍력)')

# 세트 6
doc.add_page_break()
H('세트 6. 지역 상한·1인당 잠재량·수도권 가중치', 1)
H('분석 목적 및 설정', 2)
B([('상한: ', '지역별 PV·풍력 p_nom_max = 기술적 잠재량 × 20% (기존 설비가 더 크면 기존 유지). PyPSA-Eur와 같은 방식'),
   ('1인당 잠재량: ', '배율 지표를 잠재량 ÷ 인구(2024 주민등록인구 근사)로 변경 → 서울경기·인천·부산 배율 상승. '
    '면적당 잠재량은 인천이 가장 높게 나와(34.7 GW/천km²) 부적합'),
   ('수도권 가중치: ', '서울경기·인천의 PV·풍력 비용 ×1.3 / ×1.5, 나머지 지역은 그대로'),
   ('조합: ', '상한 20% + 1인당 잠재량 α=0.25')])
T(t6, fmt={'총비용 (조원)': '{:.2f}', '잠재량 대비 최대 지역 (%)': '{:.0f}'}, font=7.5,
  caption='표 6. 케이스별 전국 결과 (수도권 ×1.5는 ×1.3과 동일하여 생략)')
F('fig6_cap_percap_metro.png', '그림 6. 지역별 설비용량 / 기술적 잠재량 (%) — PV(a), 풍력(b)', width=17)
H('주요 결과 및 해석', 2)
B([('비용 차등만으로는 집중이 이동할 뿐: ', '1인당 잠재량 차등 시 강원 풍력이 잠재량의 91~99%, 수도권 가중치 시 충북 풍력 101%, '
    '부산 PV 19.8 GW(52%)'),
   ('수도권 가중치: ', '×1.3에서 이미 수도권 신규 설치가 0이 되어 ×1.5와 결과 동일'),
   ('상한의 효과: ', '상한 적용 케이스는 모든 지역이 잠재량의 29% 이내. 비용 차등은 "상한 내에서 어느 지역을 먼저 채우는지"를 결정'),
   ('현실화 비용: ', '상한 20% + 1인당 α=0.25 케이스는 재생 393 GW(+14 GW), 출력제한 12.8%, CO2 그림자가격 537. '
    '비용 공통 상한 케이스 기준 입지 현실화 비용은 약 0.4~0.5조원/년'),
   ('정규화 주의: ', '1인당 α=0.25/0.5 단독 케이스의 총비용 감소(56.2/51.2조원)는 신규 평균 배율(약 0.85/0.70) 하락 효과')])
H('시사점 (시나리오 반영)', 2)
B(['지역 입지 현실화는 "잠재량 기반 상한 + 1인당(또는 도시화) 지표 기반 비용 차등"의 2단 구조가 적절',
   '상한 비율(20%)은 시장잠재량·정책목표와 비교하여 지역별로 조정 가능한 시나리오 변수로 운용 권장',
   '입지를 현실화하면 재생 용량·출력제한이 증가하므로, 송전망(세트 7)과 유연성(세트 2~4) 시나리오를 함께 제시해야 함'])

# 세트 7
doc.add_page_break()
H('세트 7. 송전망 확장', 1)
H('분석 목적 및 설정', 2)
B(['기준: 세트 6의 "상한 20% + 1인당 α=0.25" 케이스 (현재 입력에서 모든 송전선·HVDC는 확장 불가)',
   ('필요 구간 확장: ', '구간별 양방향 증설 링크(확장 가능) 추가 + HVDC 3회선 확장 허용 → 최적화가 증설 구간·규모 결정. '
    '건설비 ×0.5 / ×1 / ×2'),
   ('일괄 확장: ', '전 AC 구간·HVDC 용량 +10% / +20% / +30% 고정, 증설 비용은 동일 단가로 사후 합산'),
   ('건설비 가정: ', 'AC 345kV 2,000 천원/MW·km(최소 30 km), HVDC 3,000 천원/MW·km + 변환소 300,000 천원/MW, '
    '40년·4.5%·운영유지비 1%/년 연간화(계수 0.064)')])
T(t7, fmt={'송전 연간비용 (조원)': '{:.2f}', '총비용·송전포함 (조원)': '{:.2f}'}, font=7.5,
  caption='표 7-1. 송전망 확장 방식별 전국 결과')
F('fig7_transmission.png', '그림 7. 필요 구간별 증설량(a)과 송전 증설량–총비용 관계(b)', width=17)
t7s = t7c[(t7c[['필요구간 (건설비×1)', '필요구간 (건설비×0.5)', '필요구간 (건설비×2)']].max(axis=1) > 0.1)]
T(t7s[['유형', '구간', '기존 (GW)', '필요구간 (건설비×0.5)', '필요구간 (건설비×1)', '필요구간 (건설비×2)',
       '혼잡시간 확장없음 (h)', '혼잡시간 필요구간×1 (h)']],
  fmt={'혼잡시간 확장없음 (h)': '{:,.0f}', '혼잡시간 필요구간×1 (h)': '{:,.0f}'}, font=7.5,
  caption='표 7-2. 필요 구간 확장 시 증설 구간 (GW, 그 외 AC 구간·HVDC 3회선은 증설 없음)')
H('주요 결과 및 해석', 2)
B([('필요 구간: ', '건설비 ×1에서 5개 구간 35 GW 증설(강원–대구경북 +20 GW, 서울경기–충남 +7 GW, 경남–전북 +6 GW 등). '
    '건설비 ×0.5에서는 대구경북–충북(+2 GW)이 추가되고, ×2에서는 경남–광주전남이 사실상 빠지지만 핵심 3개 구간은 공통'),
   ('효율성: ', '필요 구간 35 GW 증설(연 0.16조원)로 총비용 −1.5조원. 일괄 +20%는 39 GW에 0.94조원이 들고 총비용은 증가'),
   ('HVDC: ', '증설되지 않음. 대신 HVDC로 전력을 보내는 AC 경로를 보강해 동해안 HVDC 혼잡시간 1,609→79시간'),
   ('용량·출력제한: ', '재생 393→387 GW(−5.5 GW), 출력제한 12.8→12.3%로 감소폭 제한적 → 출력제한의 주원인은 봄철 전국 동시 잉여'),
   ('CO2: ', '그림자가격 537→384, 가스 발전 160.9→158.6 TWh')])
H('시사점 (시나리오 반영)', 2)
B(['송전망 시나리오는 "일괄 비율 증설"보다 "병목 구간 집중 증설"로 설계하는 것이 비용 효율적',
   '강원–대구경북, 서울경기–충남, 영·호남 연계(경남–전북·광주전남)가 우선 보강 후보',
   '송전망은 재생 용량보다 총비용·CO2 감축 부담 완화 효과가 크므로, 보고서에서는 비용 지표 중심으로 제시',
   '모형 한계: 신규 용량은 조류 제어 가능한 링크로 근사(KVL 미적용)하여 우회 경로 효과가 다소 과대평가될 수 있음. '
   '실제 선로 단가·경과지 제약 반영 시 재분석 권장'])

# 종합
doc.add_page_break()
H('종합 시사점 및 분석 시나리오 조정 제안', 1)
T(pd.DataFrame([
    ['출력제한 페널티 30', '363', '8.0', '+0.26', '시장 가격신호 반영'],
    ['풍력 80 GW 고정 + ESS 26.7%', '364', '9.3', '−0.37', '풍력 목표 유지 + ESS 확대'],
    ['페널티 30 + ESS 38.2%', '365', '7.9', '+0.22', '가격신호 + ESS 손익분기'],
    ['원자력 0.6 + 페널티 30', '364', '6.4', '+0.15', '원자력 유연운전 + 가격신호'],
    ['상한 20% + 1인당 α0.25', '393', '12.8', '+0.41', '지역 입지 현실화'],
    ['위 + 필요 구간 송전 확장', '387', '12.3', '−1.12', '입지 현실화 + 송전 보강'],
], columns=['시나리오 후보', '재생 (GW)', '출력제한 (%)', '총비용 변화 (조원/년, 기본 대비)', '정책 메시지']), font=8,
  caption='표 8. 시나리오 후보 비교 (기본: 379 GW, 11.4%, 59.80조원)')
B([('1. 재생 보급 규모: ', 'CO2 제약 하에서 재생 발전량은 고정되므로, 보급 용량은 PV/풍력 구성비 가정이 좌우한다. '
    '풍력 보급목표를 시나리오 하한으로 명시할 것을 권장'),
   ('2. 유연성 시나리오: ', 'ESS·V2G·원자력 유연성은 비용과 출력제한률을 낮추는 수단으로 제시하고, 용량 감축 효과는 '
    '가격신호·풍력 목표와 결합할 때만 나타남을 명시'),
   ('3. 지역 입지: ', '잠재량 기반 상한(예: 20%)을 기본 제약으로 두고, 1인당 잠재량 등 도시화 지표로 비용 차등을 보조 적용'),
   ('4. 송전망: ', '입지 현실화 시나리오에는 병목 구간 보강을 함께 반영. 일괄 증설은 비효율'),
   ('5. 지표 제시: ', '출력제한률 단독 제시는 오해 소지(원자력 감발 대체 등). 재생 용량·총비용·CO2 그림자가격을 함께 제시')])
H('추가 검토 과제', 2)
B(['실제 송전 건설 단가·경과지 제약 반영, AC 선로 확장(KVL 포함) 재시도(솔버 설정 조정)',
   '상한 비율의 지역별 차등(시장잠재량·정책목표 연계), 도시지역 면적 기반 가용 잠재량 산정',
   '열 부문 보강 발전(heat fallback, 약 1.5 TWh·7.3조원)에 대한 히트펌프 확장 허용 검토',
   '설비 적정 예비율(용량가치) 제약 반영 시 ESS·풍력 가치 변화 검토'])

# 부록
doc.add_page_break()
H('부록. 방법론 상세 및 결과 파일', 1)
H('A. 수식 및 구현', 2)
B([('출력제한 페널티: ', 'capital_cost_g += c × Σ_t p_max_pu_g,t × w_t (확장 가능 PV·풍력), marginal_cost_g −= c → '
    '목적함수 증분 = c × Σ(가용 − 발전) = c × 출력제한량. 실제 총비용은 페널티 항을 제외하여 재계산'),
   ('지역 비용 배율: ', 'm_r = k (P_r / mean P)^(−α), k = Σ(기존용량 × raw) / Σ기존용량. 1인당 지표는 P_r / 인구_r'),
   ('지역 상한: ', 'p_nom_max = max(기술적 잠재량 × 20%, 기존 용량)'),
   ('송전 연간화: ', 'CRF(40년, 4.5%) 0.0543 + 운영유지비 0.01 = 0.0643'),
   ('솔버: ', 'CPLEX Barrier convergetol 1e-6(세트 1~5), 1e-5(세트 6 일부·세트 7: 수렴 실패 케이스 대응, 환경변수 '
    'PYPSA_CPLEX_CONVERGETOL)')])
H('B. 결과 파일', 2)
T(pd.DataFrame([
    ['1', r'results\ESS_cost_sweep_20261028_095343', 'run_ess_cost_sweep.py'],
    ['2·3', r'results\flex_sweep_20261028_113555', 'run_flex_sweep.py --set AB'],
    ['4', r'results\flex_sweep_CDE_20261028_142657', 'run_flex_sweep.py --set CDE'],
    ['5', r'results\regional_cost_sweep_20261028_154525', 'run_regional_cost_sweep.py'],
    ['6', r'results\regional_cost_sweep_mix_20261028_161435, ..._171926', 'run_regional_cost_sweep.py --set mix'],
    ['7', r'results\transmission_sweep_20261028_224454, ..._231818, 20261029_111446', 'run_transmission_sweep.py'],
], columns=['세트', '결과 폴더', '실행 스크립트']), font=8)
H('C. 유의사항', 2)
B(['모든 결과는 단일 기준연도·현재 입력조건 기준이며, 수요·비용 전망 변경 시 수치는 달라질 수 있음',
   '총비용은 기존 설비 고정비를 제외한 값으로, 케이스 간 "차이"로 해석해야 함',
   '비용 배율을 바꾼 케이스(세트 1, 3, 5, 6)의 총비용은 바뀐 비용 가정 하의 값이므로 절대 비교 시 주의',
   '인구 자료는 2024년 주민등록인구 근사치로, 보고서 활용 시 공식 통계로 갱신 권장'])

docx_path = os.path.join(OUT, '민감도분석_보고서.docx')
doc.save(docx_path)

xlsx_path = os.path.join(OUT, '민감도분석_표.xlsx')
with pd.ExcelWriter(xlsx_path) as w:
    for name, df in excel_tables.items():
        df.to_excel(w, sheet_name=name[:31], index=False)
print('완료:', docx_path, xlsx_path, sep='\n')
