"""compare_wp2.py 결과(CSV)로 추가 지표 계산 및 그림 생성"""
import glob
import os

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from run_regional_cost_sweep import TECH_POTENTIAL_GW

plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 9

D = sorted(glob.glob('reports/wp2_compare_*'))[-1]
nat = pd.read_csv(f'{D}/national.csv', index_col=0)
reg = pd.read_csv(f'{D}/regional.csv')
corr = pd.read_csv(f'{D}/corridor.csv')
REGION_KR = {'BSN': '부산울산', 'CBD': '충북', 'CND': '대전세종충남', 'GBD': '대구경북', 'GGD': '서울경기',
             'GND': '경남', 'GWD': '강원', 'ICN': '인천', 'JBD': '전북', 'JJD': '제주', 'JND': '광주전남'}
REG = list(REGION_KR)
PAIRS = [('2035', 'OPEN_2035', 'WP2fix_2035'), ('2040', 'OPEN_2040', 'WP2_2040')]
pd.set_option('display.width', 250)

# 추가 지표
ext = {}
for s in nat.columns:
    r = reg[reg.scenario == s].set_index('region')
    pot_use = {}
    for tech, col in (('PV', 'PV_GW'), ('WT', 'WT_GW')):
        u = r[col] / TECH_POTENTIAL_GW[tech].reindex(r.index) * 100
        pot_use[tech] = u
    fab = r.Fab_TWh
    ext[s] = {
        '수도권 순수입 (TWh)': -(r.loc[['GGD', 'ICN'], 'net_export_TWh'].sum()),
        '수도권 수요 비중 (%)': 100 * r.loc[['GGD', 'ICN'], 'load_TWh'].sum() / r.load_TWh.sum(),
        '반도체 가중 재생 자급률 (%)': (fab * r['RE_selfsuff_%']).sum() / fab.sum(),
        '반도체 가중 전력가격 (천원/MWh)': (fab * r.price_loadw).sum() / fab.sum(),
        '잠재량 초과 지역·기술 수': int((pot_use['PV'] > 100).sum() + (pot_use['WT'] > 100).sum()),
        '충북 풍력 잠재량 대비 (%)': pot_use['WT']['CBD'],
        '서울경기 풍력 잠재량 대비 (%)': pot_use['WT']['GGD'],
        '인천 PV 잠재량 대비 (%)': pot_use['PV']['ICN'],
        'HVDC 수송량 (TWh)': corr[(corr.scenario == s) & (corr.type == 'HVDC')].flow_abs_TWh.sum(),
    }
ext = pd.DataFrame(ext)
ext.to_csv(f'{D}/extra_indicators.csv', encoding='utf-8-sig')
print(ext.round(1).to_string())

# 그림 1: 반도체 수요 지역 배분
fig, ax = plt.subplots(1, 2, figsize=(10, 3.3), sharey=True)
for k, (y, a, b) in enumerate(PAIRS):
    A = reg[reg.scenario == a].set_index('region').Fab_TWh.reindex(REG)
    B = reg[reg.scenario == b].set_index('region').Fab_TWh.reindex(REG)
    x = np.arange(len(REG))
    ax[k].bar(x - 0.2, A, 0.4, color='#9E9E9E', label='기준(정부 발표)')
    ax[k].bar(x + 0.2, B, 0.4, color='#8E5EA2', label='재배치(기업 설문)')
    ax[k].set_xticks(x, [REGION_KR[r] for r in REG], rotation=45, ha='right', fontsize=8)
    ax[k].set_title(f'({chr(97 + k)}) {y}년 반도체 수요 (TWh)'); ax[k].legend(fontsize=8); ax[k].grid(axis='y', alpha=0.3)
fig.tight_layout(); fig.savefig(f'{D}/fig1_fab_allocation.png', dpi=200); plt.close(fig)

# 그림 2: 지역별 변화 (재생설비, 순수입)
fig, ax = plt.subplots(1, 3, figsize=(12, 3.6), sharey=True)
for k, (y, a, b) in enumerate(PAIRS):
    A = reg[reg.scenario == a].set_index('region').reindex(REG)
    B = reg[reg.scenario == b].set_index('region').reindex(REG)
    yy = np.arange(len(REG))[::-1]
    ax[k].barh(yy + 0.2, B.PV_GW - A.PV_GW, 0.4, color='#F2B134', label='PV')
    ax[k].barh(yy - 0.2, B.WT_GW - A.WT_GW, 0.4, color='#3C8DBC', label='풍력')
    ax[k].axvline(0, color='k', lw=0.6)
    ax[k].set_yticks(yy, [REGION_KR[r] for r in REG], fontsize=8)
    ax[k].set_xlabel('기준 대비 변화 (GW)'); ax[k].set_title(f'({chr(97 + k)}) {y}년 재생설비 변화')
    ax[k].legend(fontsize=7.5, loc='lower right'); ax[k].grid(axis='x', alpha=0.3)
for j, (y, a, b) in enumerate(PAIRS):
    A = reg[reg.scenario == a].set_index('region').reindex(REG)
    B = reg[reg.scenario == b].set_index('region').reindex(REG)
    yy = np.arange(len(REG))[::-1]
    ax[2].barh(yy + (0.2 if j == 0 else -0.2), -(B.net_export_TWh - A.net_export_TWh), 0.4,
               color='#7FA7C9' if j == 0 else '#1F4E79', label=f'{y}년')
ax[2].axvline(0, color='k', lw=0.6)
ax[2].set_xlabel('순수입 변화 (TWh, +는 수입 증가)'); ax[2].set_title('(c) 지역 순수입량 변화')
ax[2].legend(fontsize=7.5); ax[2].grid(axis='x', alpha=0.3)
fig.tight_layout(); fig.savefig(f'{D}/fig2_regional_change.png', dpi=200); plt.close(fig)

# 그림 3: 송전 구간별 수송량·혼잡시간 (2040)
a, b = 'OPEN_2040', 'WP2_2040'
A = corr[corr.scenario == a].set_index(['type', 'corridor'])
B = corr[corr.scenario == b].set_index(['type', 'corridor'])
d = pd.DataFrame({'base_flow': A.flow_abs_TWh, 'new_flow': B.flow_abs_TWh, 'base_h': A.cong_h, 'new_h': B.cong_h})
d = d[(d.base_flow > 0.5)]
d = d.loc[(d.new_flow - d.base_flow).abs().sort_values(ascending=True).index]
lab = [f"{'[HVDC] ' if t == 'HVDC' else ''}{REGION_KR[c.split('-')[0]]}–{REGION_KR[c.split('-')[1]]}" for t, c in d.index]
fig, ax = plt.subplots(1, 2, figsize=(11, 4.6), sharey=True)
yy = np.arange(len(d))
ax[0].barh(yy + 0.2, d.base_flow, 0.4, color='#9E9E9E', label='기준')
ax[0].barh(yy - 0.2, d.new_flow, 0.4, color='#1F4E79', label='재배치')
ax[0].set_yticks(yy, lab, fontsize=8); ax[0].set_xlabel('연간 수송량 (TWh, 양방향 합)')
ax[0].set_title('(a) 2040년 구간별 송전 수송량'); ax[0].legend(fontsize=8); ax[0].grid(axis='x', alpha=0.3)
ax[1].barh(yy + 0.2, d.base_h, 0.4, color='#9E9E9E', label='기준')
ax[1].barh(yy - 0.2, d.new_h, 0.4, color='#C0392B', label='재배치')
ax[1].set_xlabel('혼잡시간 (h/년)'); ax[1].set_title('(b) 2040년 구간별 혼잡시간'); ax[1].legend(fontsize=8)
ax[1].grid(axis='x', alpha=0.3)
fig.tight_layout(); fig.savefig(f'{D}/fig3_grid_2040.png', dpi=200); plt.close(fig)

# 그림 4: 상한 설정(REF) vs 해제(OPEN) — 지역별 slack과 재생설비 (2040)
r1 = reg[reg.scenario == 'REF_2040'].set_index('region').reindex(REG)
r2 = reg[reg.scenario == 'OPEN_2040'].set_index('region').reindex(REG)
fig, ax = plt.subplots(1, 2, figsize=(10, 3.4))
x = np.arange(len(REG))
ax[0].bar(x - 0.2, r1.PV_GW + r1.WT_GW, 0.4, color='#BDBDBD', label='REF(상한)')
ax[0].bar(x + 0.2, r2.PV_GW + r2.WT_GW, 0.4, color='#2E7D32', label='상한 해제')
ax[0].set_xticks(x, [REGION_KR[r] for r in REG], rotation=45, ha='right', fontsize=8)
ax[0].set_ylabel('PV+풍력 (GW)'); ax[0].set_title('(a) 2040년 지역별 재생설비'); ax[0].legend(fontsize=8)
ax[1].bar(x, r1.slack_TWh, color='#C0392B')
for i, v in enumerate(r1.slack_TWh):
    if v > 0.5:
        ax[1].text(i, v + 1, f'{v:.0f}', ha='center', fontsize=8)
ax[1].set_xticks(x, [REGION_KR[r] for r in REG], rotation=45, ha='right', fontsize=8)
ax[1].set_ylabel('TWh'); ax[1].set_title('(b) 2040년 REF의 지역별 공급부족(slack)')
fig.tight_layout(); fig.savefig(f'{D}/fig4_ref_vs_open_2040.png', dpi=200); plt.close(fig)
print('figs done', D)
