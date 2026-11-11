"""
대표경로 2035년(A·B 시나리오) 분석 보고서(Word) + 그림 + 표 생성

선행:
  python run_2035_sensitivity.py --only base A
  python run_mga_sweep.py --tag path2035B --potential-cap 0.2 --demand DC=0.5 Fab=0.916 --eps 1 --no-penalty --barrier-alg 1
  python analyze_2035.py <민감도 결과폴더> <MGA 결과폴더> --out reports/path2035_<날짜>
사용법:
  python make_path2035_report.py --sens-dir results/sens2035_... --mga-dir results/mga_sweep_path2035B_... [--dir reports/path2035_...]
"""
import argparse
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

from analyze_mga import kind
from run_regional_cost_sweep import TECH_POTENTIAL_GW

logging.disable(logging.WARNING)
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 9

ap = argparse.ArgumentParser()
ap.add_argument('--dir', default=sorted(glob.glob('reports/path2035_*'))[-1])
ap.add_argument('--sens-dir', required=True)
ap.add_argument('--mga-dir', required=True)
ap.add_argument('--capacity-dir', default='results/path2035_capacity')
args = ap.parse_args()
OUT = args.dir

nat = pd.read_csv(f'{OUT}/national.csv', index_col=0)
reg = pd.read_csv(f'{OUT}/regional.csv')
NET = {'base': f'{args.sens_dir}/base.nc', 'A': f'{args.sens_dir}/A.nc', 'B': f'{args.mga_dir}/MGA-1.nc'}
nat = nat.rename(index={'MGA-1': 'B'})
reg['case'] = reg.case.replace({'MGA-1': 'B'})

REGION_KR = {'BSN': '부산울산', 'CBD': '충북', 'CND': '대전세종충남', 'GBD': '대구경북', 'GGD': '서울경기',
             'GND': '경남', 'GWD': '강원', 'ICN': '인천', 'JBD': '전북', 'JJD': '제주', 'JND': '광주전남'}
REG = ['GGD', 'ICN', 'GWD', 'CBD', 'CND', 'JBD', 'GBD', 'BSN', 'GND', 'JND', 'JJD']
CASE_KR = {'base': '기준', 'A': 'A', 'B0': 'B0 (MGA 기준)', 'B': 'B'}
CASES = ['base', 'A', 'B']
C_PV, C_WT, C_B, C_PEN, C_GAS, C_NU, C_CHP = '#F2B134', '#3C8DBC', '#1F4E79', '#C0392B', '#E07B39', '#6C757D', '#8E5EA2'
COLS = {'base': '#BDD7EE', 'A': '#2E75B6', 'B': '#1F4E79'}

R = {c: reg[reg.case == c].set_index('region').reindex(REG) for c in nat.index}
LIM = {t: {r: max(0.2 * TECH_POTENTIAL_GW[t][r], float(R['base'].at[r, f'{t}_exist_GW'])) for r in REG} for t in ('PV', 'WT')}


def V(c, k):
    return float(nat.at[c, k])


def d(c, k, ref):
    return V(c, k) - V(ref, k)


def pct(c, ref):
    return 100 * (V(c, 'cost_sys') / V(ref, 'cost_sys') - 1)


figs = {}


def save(fig, name, rect=(0, 0, 1, 1)):
    p = os.path.join(OUT, name)
    fig.tight_layout(rect=rect)
    fig.savefig(p, dpi=200)
    plt.close(fig)
    figs[name] = p


# 승계 대상 설비 (확장 가능 설비의 최적 용량)
def build_summary(path):
    n = pypsa.Network(path)
    rows = {}
    for comp, attr in (('Generator', 'p_nom'), ('Link', 'p_nom'), ('Store', 'e_nom'), ('Line', 's_nom')):
        df = n.static(comp)
        ext = df.index[df[f'{attr}_extendable'].astype(bool)]
        for nm in ext:
            k = {'solar': 'PV', 'wind': 'WT'}.get(df.at[nm, 'carrier'], kind(nm)) if comp == 'Generator' else \
                ('AC 선로' if comp == 'Line' else kind(nm))
            unit = 'GWh' if comp == 'Store' else 'GW'
            key = (k, unit)
            e, o = float(df.at[nm, attr]), float(df.at[nm, f'{attr}_opt'])
            r = rows.setdefault(key, [0.0, 0.0])
            r[0] += e / 1e3; r[1] += o / 1e3
    return rows


BUILD = {c: build_summary(p) for c, p in NET.items()}
keys = sorted({k for b in BUILD.values() for k, v in b.items() if any(abs(BUILD[c].get(k, [0, 0])[1] - BUILD[c].get(k, [0, 0])[0]) > 0.05
                                                                     for c in CASES)})
KEY_KR = {'PV': '태양광', 'WT': '풍력', 'ESS': 'ESS 저장용량', 'ESS_Ch': 'ESS 충전', 'ESS_Disch': 'ESS 방전',
          'Electrolyser': '수전해', 'Fuelcell': '연료전지', 'H2_Storage': '수소 저장', 'H_Storage': '열 저장',
          'Heat_Ch': '열저장 충전', 'Heat_Disch': '열저장 방전', 'HVDC': 'HVDC', 'AC 선로': 'AC 선로', 'slack': '전력 slack 설비'}
build_tbl = pd.DataFrame([{'설비': f"{KEY_KR.get(k, k)} ({u})", '기존 (입력)': BUILD['base'][(k, u)][0],
                           **{f'{CASE_KR[c]} 최적': BUILD[c].get((k, u), [0, 0])[1] for c in CASES},
                           **{f'{CASE_KR[c]} 신규': BUILD[c].get((k, u), [0, 0])[1] - BUILD[c].get((k, u), [0, 0])[0] for c in ('A', 'B')}}
                          for k, u in keys if k not in ('slack', 'H_Fallback_Gen', 'LNG_Fuel_Supply')])

# ───────────── 그림 ─────────────
fig, ax = plt.subplots(1, 2, figsize=(11, 3.8))
xx = np.arange(len(CASES))
pv = [V(c, 'PV_GW') for c in CASES]; wt = [V(c, 'WT_GW') for c in CASES]
ax[0].bar(xx, pv, 0.55, color=C_PV, label='PV'); ax[0].bar(xx, wt, 0.55, bottom=pv, color=C_WT, label='풍력')
for i in range(len(CASES)):
    ax[0].text(i, pv[i] / 2, f'{pv[i]:.1f}', ha='center', fontsize=8)
    ax[0].text(i, pv[i] + wt[i] / 2, f'{wt[i]:.1f}', ha='center', fontsize=8, color='white')
    ax[0].text(i, pv[i] + wt[i] + 4, f'{pv[i] + wt[i]:.1f}', ha='center', fontsize=8.5, fontweight='bold')
ax[0].set_xticks(xx, [CASE_KR[c] for c in CASES]); ax[0].set_ylabel('GW'); ax[0].set_ylim(0, 320)
ax[0].legend(fontsize=8, loc='lower right'); ax[0].grid(axis='y', alpha=0.3); ax[0].set_title('(a) PV·풍력 설비')
cs = [V(c, 'cost_sys') for c in CASES]
ax[1].bar(xx, cs, 0.55, color=[COLS[c] for c in CASES])
for i, c in enumerate(CASES):
    ax[1].text(i, cs[i] + 0.8, f'{cs[i]:.2f}', ha='center', fontsize=8.5)
ax2 = ax[1].twinx()
ax2.plot(xx, [V(c, 'curtail_%') for c in CASES], '-o', color=C_PEN, label='출력제한률')
ax2.set_ylim(0, 14); ax2.set_ylabel('출력제한률 (%)'); ax2.legend(fontsize=8, loc='upper right')
ax[1].set_xticks(xx, [CASE_KR[c] for c in CASES]); ax[1].set_ylabel('시스템 비용 (조원/년)'); ax[1].set_ylim(0, 75)
ax[1].set_title('(b) 시스템 비용(열보조 페널티 제외)과 출력제한률'); ax[1].grid(axis='y', alpha=0.3)
save(fig, 'fig1_overview.png')

fig, ax = plt.subplots(1, 2, figsize=(12, 4.2), sharey=True)
yy = np.arange(len(REG))[::-1]
for k, t in enumerate(('PV', 'WT')):
    for j, c in enumerate(CASES):
        ax[k].barh(yy + 0.27 - 0.27 * j, R[c][f'{t}_GW'], 0.27, color=COLS[c], label=CASE_KR[c])
    ax[k].scatter([LIM[t][r] for r in REG], yy, marker='|', s=260, color=C_PEN, zorder=4, label='상한 (잠재량 20%)')
    ax[k].scatter(R['base'][f'{t}_exist_GW'], yy, marker='x', s=20, color='k', zorder=4, label='기존 설비')
    ax[k].set_title(f"({'ab'[k]}) {'태양광' if t == 'PV' else '풍력'} 지역별 설비 (GW)"); ax[k].grid(axis='x', alpha=0.3)
ax[0].set_yticks(yy, [REGION_KR[r] for r in REG])
h, l_ = ax[0].get_legend_handles_labels()
fig.legend(h, l_, fontsize=8, loc='lower center', ncol=5, bbox_to_anchor=(0.5, -0.01))
save(fig, 'fig2_regional.png', rect=(0, 0.07, 1, 1))

items = [('PV 발전', 'PV_TWh'), ('풍력 발전', 'WT_TWh'), ('출력제한', 'curtail_TWh'), ('원자력', 'nuclear_TWh'),
         ('LNG 발전', 'LNG_power_TWh'), ('CHP 발전', 'CHP_power_TWh'), ('석탄', 'coal_TWh'),
         ('수전해 전력', 'electrolyser_elec_TWh'), ('EV 충전', 'EV_charge_TWh'), ('ESS 충전', 'ESS_charge_TWh'),
         ('일반·DC·Fab 수요', 'demand_power_TWh')]
fig, ax = plt.subplots(figsize=(11, 3.8))
xx = np.arange(len(items))
ax.bar(xx - 0.2, [d('A', k, 'base') for _, k in items], 0.4, color=COLS['A'], label='A - 기준')
ax.bar(xx + 0.2, [d('B', k, 'A') for _, k in items], 0.4, color=C_B, label='B - A')
ax.axhline(0, color='k', lw=0.6)
ax.set_xticks(xx, [n_ for n_, _ in items], rotation=25, ha='right', fontsize=8)
ax.set_ylabel('연간 변화 (TWh)'); ax.legend(fontsize=8); ax.grid(axis='y', alpha=0.3)
ax.set_title('시나리오 간 연간 전력 수급 변화')
save(fig, 'fig3_balance.png')

groups = [('PV 투자비', ['c_capex_PV']), ('풍력 투자비', ['c_capex_WT']), ('가스 연료비', ['c_opex_gas']),
          ('CHP 운전비', ['c_opex_CHP']), ('원자력·석탄 연료비', ['c_opex_nuclear', 'c_opex_coal']),
          ('수소·연료전지', ['c_capex_Electrolyser', 'c_capex_Fuelcell', 'c_opex_Fuelcell']),
          ('저장·V2G·기타', ['c_capex_ESS', 'c_capex_H_Storage', 'c_opex_EV_V2G', 'c_opex_HP', 'c_opex_DR',
                          'c_opex_ESS_Ch', 'c_opex_ESS_Disch', 'c_capex_ESS_Ch', 'c_capex_ESS_Disch'])]
gcol = [C_PV, C_WT, C_GAS, C_CHP, C_NU, '#2E7D32', '#9E9E9E']


def gsum(c, ks):
    return sum(0 if (k not in nat or pd.isna(nat.at[c, k])) else float(nat.at[c, k]) for k in ks)


fig, ax = plt.subplots(figsize=(9.5, 3.8))
xx = np.arange(len(CASES)); bot = np.zeros(len(CASES))
for (nm, ks), col in zip(groups, gcol):
    vals = np.array([gsum(c, ks) for c in CASES])
    ax.bar(xx, vals, 0.55, bottom=bot, color=col, label=nm)
    bot += vals
for i, c in enumerate(CASES):
    ax.text(i, V(c, 'cost_sys') + 0.8, f"{V(c, 'cost_sys'):.2f}", ha='center', fontsize=8.5)
ax.set_xticks(xx, [CASE_KR[c] for c in CASES]); ax.set_ylabel('조원/년'); ax.set_ylim(0, 72)
ax.legend(fontsize=7.5, loc='center left', bbox_to_anchor=(1.01, 0.5)); ax.grid(axis='y', alpha=0.3)
ax.set_title('시스템 비용 구성 (열보조 페널티 제외)')
save(fig, 'fig4_cost.png')


# ───────────── 표 ─────────────
def summary_table():
    rows = []
    for c in CASES:
        rows.append({'시나리오': CASE_KR[c], '전력수요 (TWh)': V(c, 'demand_power_TWh'), 'PV (GW)': V(c, 'PV_GW'),
                     '풍력 (GW)': V(c, 'WT_GW'), 'PV+풍력 (GW)': V(c, 'RE_GW'), '재생 발전 (TWh)': V(c, 'PV_TWh') + V(c, 'WT_TWh'),
                     '출력제한률 (%)': V(c, 'curtail_%'), 'ESS (GW/GWh)': f"{V(c, 'ESS_GW'):.1f} / {V(c, 'ESS_GWh'):.1f}",
                     '시스템 비용 (조원)': V(c, 'cost_sys'), 'CO2 잠재가격 (천원/t)': V(c, 'CO2_price'),
                     '부하가중 전력가격 (천원/MWh)': V(c, 'EL_price_loadw')})
    return pd.DataFrame(rows)


def mga_table():
    rows = []
    for c in ('B0', 'B'):
        rows.append({'시나리오': CASE_KR[c], '비용 증가 허용률': '—' if c == 'B0' else '1%',
                     '시스템 비용 (조원)': V(c, 'cost_sys'), '실제 비용 증가율 (%)': 0.0 if c == 'B0' else pct(c, 'B0'),
                     '태양광 GW': V(c, 'PV_GW'), '풍력 GW': V(c, 'WT_GW'), 'B0 대비 재생 감소 GW': V('B0', 'RE_GW') - V(c, 'RE_GW'),
                     'ESS GW/GWh': f"{V(c, 'ESS_GW'):.1f} / {V(c, 'ESS_GWh'):.1f}", '출력제한율 (%)': V(c, 'curtail_%')})
    return pd.DataFrame(rows)


def balance_table():
    its = [('일반 수요', 'demand_EL_TWh'), ('DC 수요', 'demand_DC_TWh'), ('Fab 수요', 'demand_Fab_TWh'),
           ('PV 발전', 'PV_TWh'), ('풍력 발전', 'WT_TWh'), ('출력제한', 'curtail_TWh'), ('원자력', 'nuclear_TWh'),
           ('석탄', 'coal_TWh'), ('LNG 발전', 'LNG_power_TWh'), ('CHP 발전', 'CHP_power_TWh'), ('CHP 열', 'CHP_heat_TWh'),
           ('가스 연료 총량', 'gas_fuel_TWh'), ('연료전지 발전', 'fuelcell_power_TWh'), ('히트펌프 전력', 'HP_elec_TWh'),
           ('히트펌프 열', 'HP_heat_TWh'), ('열보조공급', 'heat_fallback_TWh'), ('수전해 전력', 'electrolyser_elec_TWh'),
           ('EV 충전', 'EV_charge_TWh'), ('V2G 방전', 'V2G_TWh'), ('ESS 충전', 'ESS_charge_TWh'),
           ('전력 소비 합계', 'elec_demand_total_TWh'), ('HVDC 수송량', 'HVDC_TWh'), ('CO2 배출 (Mt)', 'CO2_Mt')]
    df = pd.DataFrame({'항목 (TWh)': [a for a, _ in its]})
    for c in CASES:
        df[CASE_KR[c]] = [V(c, k) for _, k in its]
    df['A − 기준'] = df['A'] - df['기준']; df['B − A'] = df['B'] - df['A']
    return df


def cost_table():
    df = pd.DataFrame({'비용 항목 (조원/년)': [g for g, _ in groups] + ['시스템 비용 합계', '(참고) 열보조공급 페널티']})
    for c in CASES:
        df[CASE_KR[c]] = [gsum(c, ks) for _, ks in groups] + [V(c, 'cost_sys'), V(c, 'c_opex_penalty')]
    df['A − 기준'] = df['A'] - df['기준']; df['B − A'] = df['B'] - df['A']
    return df


def regional_table():
    rows = []
    for r in REG:
        row = {'지역': REGION_KR[r]}
        for t, tk in (('PV', 'PV'), ('WT', '풍력')):
            row[f'{tk} 기존'] = R['base'].at[r, f'{t}_exist_GW']
            row[f'{tk} 상한'] = LIM[t][r]
            for c in CASES:
                row[f'{tk} {CASE_KR[c]}'] = R[c].at[r, f'{t}_GW']
        rows.append(row)
    return pd.DataFrame(rows)


TBL = {'표1_요약': summary_table(), '표2_MGA': mga_table(), '표3_수급': balance_table(), '표4_비용': cost_table(),
       '표5_지역별': regional_table(), '표6_승계설비': build_tbl}
with pd.ExcelWriter(os.path.join(OUT, '대표경로_2035_분석_표.xlsx')) as xw:
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


FM = {'시스템 비용 (조원)': '{:,.2f}', 'CO2 잠재가격 (천원/t)': '{:,.0f}', '부하가중 전력가격 (천원/MWh)': '{:,.0f}',
      '실제 비용 증가율 (%)': '{:,.2f}'}
FMC = {c: '{:,.2f}' for c in ['기준', 'A', 'B', 'A − 기준', 'B − A']}


def moves(c, ref, t, sign):
    vals = [(R[c].at[r, f'{t}_GW'] - R[ref].at[r, f'{t}_GW'], r) for r in REG]
    vals = [(v, r) for v, r in vals if sign * v > 0.5]
    return ', '.join(f"{REGION_KR[r]} {v:+.1f}" for v, r in sorted(vals, key=lambda x: -abs(x[0])))


fb_reg = {r: 0.0 for r in REG}
nb = pypsa.Network(NET['A'])
fbP = nb.generators_t.p.mul(nb.snapshot_weightings.generators, axis=0).sum() / 1e6
for g_ in nb.generators.index[nb.generators.index.str.contains('H_Fallback')]:
    fb_reg[g_[:3]] = fb_reg.get(g_[:3], 0.0) + float(fbP[g_])
fb_txt = ', '.join(f"{REGION_KR[r]} {v:.1f}" for r, v in sorted(fb_reg.items(), key=lambda x: -x[1]) if v > 0.05)
fc = nb.links.index[nb.links.index.str.endswith('_Fuelcell')]
fc_exist, fc_max = nb.links.loc[fc, 'p_nom'].sum() / 1e3, nb.links.loc[fc, 'p_nom_max'].sum() / 1e3

doc.add_heading('대표경로 2035년 분석: A·B 시나리오', 0)
P('2035년 재생에너지 보급 상한 해제 기준 + 대표 시나리오 조건 (2040년 승계용)', size=12)
P(f'작성일: {datetime.now():%Y-%m-%d}   |   모형: PyPSA 0.31.1 + linopy + CPLEX Barrier, 8,784시간 단일연도', size=9)

H('개요', 1)
P('1차 분석은 2035년과 2040년을 각각 독립적으로 계산하였다. 대표경로 분석은 2035년 결과를 2040년에 승계해 연속성 있는 경로를 만드는 것이 목적이며, '
  '이 보고서는 그 첫 단계인 2035년 A·B 시나리오 결과이다. 입력은 현재 interface 값을 그대로 쓰고, 시나리오 조건은 네트워크 생성 후 메모리에서만 적용하였다.')
T(pd.DataFrame([
    ['기준', '현재 interface 2035년 입력, 재생에너지 보급 상한 해제 (GUI 실행 결과 results/대표경로/1104/2035_A1과 목적함수·설비 일치 확인)'],
    ['A', 'DC 수요 −50% + Fab 수요 −8.4% + 지역 입지 제약(PV·풍력 상한 = max(기술적 잠재량 × 20%, 기존 설비)), 비용 최소화'],
    ['B', 'A와 같은 조건에서 비용을 A의 비용최소해보다 1% 더 허용하는 MGA. PV·풍력 신규 용량 최소화 후 그 용량 아래 비용 최소화'],
], columns=['시나리오', '조건']), font=8, caption='표 0-1. 시나리오 정의')

H('핵심 결과', 2)
B([('A (대표 조건): ', f"PV {V('A', 'PV_GW'):.1f} GW, 풍력 {V('A', 'WT_GW'):.1f} GW(합계 {V('A', 'RE_GW'):.1f} GW), 시스템 비용 "
    f"{V('A', 'cost_sys'):.2f}조원/년. 기준 대비 전력수요 {-d('A', 'demand_power_TWh', 'base'):.1f} TWh 감소로 재생설비 "
    f"{-d('A', 'RE_GW', 'base'):.1f} GW, 비용 {-d('A', 'cost_sys', 'base'):.2f}조원이 줄었다. 감소분 대부분은 풍력"
    f"({d('A', 'WT_GW', 'base'):+.1f} GW)이다."),
   ('입지 제약의 영향: ', f"인천·부산울산 PV와 서울경기·충북 풍력이 상한까지 줄고, 서울경기·경남·대전세종충남 PV와 대전세종충남·강원·인천 풍력으로 "
    f"옮겨갔다. HVDC 수송량은 {V('base', 'HVDC_TWh'):.0f} → {V('A', 'HVDC_TWh'):.0f} TWh로 늘었다."),
   ('B (A + 비용 1%): ', f"비용 +{V('B', 'cost_sys') - V('B0', 'cost_sys'):.2f}조원(+{pct('B', 'B0'):.2f}%)으로 PV·풍력이 "
    f"{V('B0', 'RE_GW') - V('B', 'RE_GW'):.1f} GW 줄어 {V('B', 'RE_GW'):.1f} GW가 된다. PV {d('B', 'PV_GW', 'A'):+.1f} GW, "
    f"풍력 {d('B', 'WT_GW', 'A'):+.1f} GW. 출력제한률은 {V('A', 'curtail_%'):.1f}% → {V('B', 'curtail_%'):.1f}%로 낮아진다."),
   ('B의 입지 여유: ', f"잠재량 상한 때문에 풍력을 늘릴 곳이 제한적이다. B의 풍력 증가는 강원·전북·대전세종충남에서만 일어나고, "
    f"PV 감소는 광주전남({R['B'].at['JND', 'PV_GW'] - R['A'].at['JND', 'PV_GW']:+.1f} GW)이 가장 크다."),
   ('유의 사항 (현재 입력): ', f"연료전지 입력이 1차 분석 때와 달라졌다(기존 {fc_exist:.2f} GW, 확장 상한 {fc_max:.2f} GW). "
    f"그 결과 세 시나리오 모두 열보조공급(H_Fallback)이 {V('A', 'heat_fallback_TWh'):.1f} TWh 쓰인다({fb_txt} TWh). "
    '1차 2035년 기준(1.7 TWh)보다 크게 늘었다. 이 페널티(약 72.6조원)는 시스템 비용에서 제외하였다.')])
T(summary_table(), font=7, caption='표 0-2. 결과 요약', fmts=FM)
F('fig1_overview.png', '그림 1. 시나리오별 PV·풍력 설비(a)와 시스템 비용·출력제한률(b)')

doc.add_page_break()
H('1. A 시나리오: 수요 조정 + 입지 제약', 1)
B([('수요: ', f"DC {V('base', 'demand_DC_TWh'):.1f} → {V('A', 'demand_DC_TWh'):.1f} TWh, Fab {V('base', 'demand_Fab_TWh'):.1f} → "
    f"{V('A', 'demand_Fab_TWh'):.1f} TWh. 일반 수요({V('A', 'demand_EL_TWh'):.1f} TWh)와 EV·열·수소 부문은 기준과 같다."),
   ('설비: ', f"PV {d('A', 'PV_GW', 'base'):+.1f} GW, 풍력 {d('A', 'WT_GW', 'base'):+.1f} GW. 수요가 줄면 비싼 풍력부터 빠진다는 "
    '1차 분석 결과와 같다. 입지 제약으로 서울경기 풍력이 막혀 신규 풍력은 대전세종충남·강원·인천으로 옮겨갔다.'),
   ('지역 재배치: ', f"PV 증가 {moves('A', 'base', 'PV', 1)} / 감소 {moves('A', 'base', 'PV', -1)} GW. "
    f"풍력 증가 {moves('A', 'base', 'WT', 1)} / 감소 {moves('A', 'base', 'WT', -1)} GW."),
   ('수급: ', f"재생 발전량 {d('A', 'PV_TWh', 'base') + d('A', 'WT_TWh', 'base'):+.1f} TWh, LNG 발전 {d('A', 'LNG_power_TWh', 'base'):+.1f} TWh, "
    f"석탄 {d('A', 'coal_TWh', 'base'):+.1f} TWh, 수전해 전력 {d('A', 'electrolyser_elec_TWh', 'base'):+.1f} TWh. CO2는 상한 "
    f"{V('A', 'CO2_Mt'):.1f} Mt에 구속된다."),
   ('비용·가격: ', f"시스템 비용 {V('base', 'cost_sys'):.2f} → {V('A', 'cost_sys'):.2f}조원({pct('A', 'base'):+.1f}%), CO2 잠재가격 "
    f"{V('base', 'CO2_price'):.0f} → {V('A', 'CO2_price'):.0f}천원/t, 부하가중 전력가격 {V('base', 'EL_price_loadw'):.0f} → "
    f"{V('A', 'EL_price_loadw'):.0f}천원/MWh. 수요 감소 효과와 입지 제약 비용이 함께 반영된 값이다(두 효과를 따로 분리하지는 않음).")])
F('fig2_regional.png', '그림 2. 지역별 PV(a)·풍력(b) 설비 (빨간 막대: 잠재량 20% 상한, ×: 기존 설비)', width=17)

doc.add_page_break()
H('2. B 시나리오: A + 비용 1% 허용 (MGA)', 1)
B([('방법: ', '1차 분석의 MGA와 같다. B0는 A와 같은 조건에서 전력 slack을 0으로 고정한 비용최소해이다'
    f"(비용 {V('B0', 'cost_sys'):.3f}조원, A {V('A', 'cost_sys'):.3f}조원과 {pct('B0', 'A'):+.2f}% 차이, 설비는 거의 같다). "
    '1단계 C ≤ 1.01·C_B0 아래 PV·풍력 신규 용량을 최소화하고, 2단계 그 용량 아래에서 비용을 다시 최소화하였다. 열보조공급량은 지역별 B0 이하로 제한하였다.'),
   ('수치 처리: ', '1차 2035년 MGA-1에서 기본 Barrier 설정이 수렴하지 않았던 경험에 따라 MGA 단계만 barrier.algorithm 1을 적용하였다. '
    '1·2단계 모두 첫 시도(수렴 기준 10⁻⁶)에 정상 종료되었다.')])
T(mga_table(), font=8, caption='표 2-1. MGA 결과 (요청 양식)', fmts=FM)
F('fig3_balance.png', '그림 3. 시나리오 간 연간 전력 수급 변화 (A − 기준, B − A)')
B([('설비: ', f"PV {d('B', 'PV_GW', 'A'):+.1f} GW, 풍력 {d('B', 'WT_GW', 'A'):+.1f} GW. PV 감소 {moves('B', 'A', 'PV', -1)} GW, "
    f"풍력 증가 {moves('B', 'A', 'WT', 1)} GW."),
   ('메커니즘: ', f"PV 잉여가 줄어 출력제한이 {V('A', 'curtail_TWh'):.1f} → {V('B', 'curtail_TWh'):.1f} TWh로, 수전해 전력이 "
    f"{d('B', 'electrolyser_elec_TWh', 'A'):+.1f} TWh, ESS 충전이 {d('B', 'ESS_charge_TWh', 'A'):+.1f} TWh 줄어든다. "
    f"CHP 발전({d('B', 'CHP_power_TWh', 'A'):+.1f} TWh)이 늘어 히트펌프 전력({d('B', 'HP_elec_TWh', 'A'):+.1f} TWh)을 일부 대신한다. "
    '1차 분석과 같은 "PV→풍력 전환 + 잉여 순환 감소" 구조이다.'),
   ('1차 분석과 비교: ', f"1차 2035년 MGA 1%(상한 해제·수요 조정 없음)는 31.3 GW(10.5%)를 줄였다. B는 "
    f"{V('B0', 'RE_GW') - V('B', 'RE_GW'):.1f} GW({100 * (1 - V('B', 'RE_GW') / V('B0', 'RE_GW')):.1f}%)로 더 작다. 입지 제약으로 "
    '풍력 전환 여지가 줄었기 때문이다(1차 2040년 잠재량 상한 MGA와 같은 경향). 단, 1차 분석과 연료전지 입력이 다르다.'),
   ('가격 지표 주의: ', f"CO2 잠재가격 {V('A', 'CO2_price'):.0f} → {V('B', 'CO2_price'):.0f}천원/t, 전력가격 "
    f"{V('A', 'EL_price_loadw'):.0f} → {V('B', 'EL_price_loadw'):.0f}천원/MWh. 재생 용량을 묶은 데 따른 잠재가격이다.")])

doc.add_page_break()
H('3. 수급·비용 상세', 1)
T(balance_table(), font=7, caption='표 3-1. 전력·열·가스 수급 (TWh)', fmts=FMC)
F('fig4_cost.png', '그림 4. 시스템 비용 구성 (열보조 페널티 제외)')
T(cost_table(), font=7.5, caption='표 3-2. 비용 구성 (조원/년)', fmts=FMC)
T(regional_table(), font=6.5, caption='표 3-3. 지역별 PV·풍력 설비 (GW)')

doc.add_page_break()
H('4. 2040년 승계 준비', 1)
P('A·B의 2035년 최적 설비 용량을 저장하였다. 2040년 분석에서는 이 용량을 각 설비의 하한(p_nom_min)으로 넣어 2035년에 지은 설비를 유지하고, '
  '그 위에 추가 확장만 허용하는 방식(PyPSA_GUI의 apply_carryover_to_input, policy=min)을 쓸 수 있다.')
T(build_tbl, font=7.5, caption='표 4-1. 확장 가능 설비의 기존·최적 용량 (승계 대상)', fmt='{:,.2f}')
B([('저장 파일: ', f"{args.capacity_dir}/capacity_2035_A.json, capacity_2035_B.json "
    '(발전기·선로·링크·저장장치 전체의 최적 용량, PyPSA_GUI extract_capacity_carryover 형식)'),
   ('승계 시 확인할 점: ', '① 2040년 interface의 기존 설비(p_nom)와 2035년 결과의 관계(이중 계상 여부) ② 2040년 잠재량 상한과 승계 설비가 충돌하는 지역 '
    '(2035년 기존 설비가 상한을 넘는 경우 상한 = 기존 설비) ③ 설비 수명·폐지 반영 여부 ④ 연료전지 등 입력이 바뀐 설비의 처리')])

H('부록. 실행 정보', 1)
T(pd.DataFrame([
    ['기준·A', args.sens_dir, 'run_2035_sensitivity.py --only base A'],
    ['B0·B', args.mga_dir, 'run_mga_sweep.py --tag path2035B --potential-cap 0.2 --demand DC=0.5 Fab=0.916 --eps 1 --no-penalty --barrier-alg 1'],
    ['승계 용량', args.capacity_dir, 'PyPSA_GUI.extract_capacity_carryover / save_capacity_snapshot'],
    ['집계·보고서', OUT, 'analyze_2035.py, make_path2035_report.py'],
], columns=['내용', '폴더', '실행']), font=7.5)
B([('비용 정의: ', 'PyPSA 목적함수(신규 설비 연간화 투자비 + 전체 운전비, 기존 설비 고정비 제외) − 열보조공급 페널티.'),
   ('기준 검증: ', 'interface 현재 입력으로 다시 계산한 기준의 목적함수(133.19조원)와 PV·풍력 설비가 GUI 결과(2035_A1)와 일치.'),
   ('1차 분석과 입력 차이: ', '연료전지 기존 설비·확장 상한만 다르다. 수요·CO2 상한·나머지 설비는 같다.')])

path = os.path.join(OUT, '대표경로_2035_AB_분석_보고서.docx')
doc.save(path)
print(f'완료: {path}')
pd.set_option('display.width', 250)
print(summary_table().round(2).to_string(index=False))
print(mga_table().round(2).to_string(index=False))
print(build_tbl.round(2).to_string(index=False))
print(cost_table().round(2).to_string(index=False))
print(fb_txt, fc_exist, fc_max)
