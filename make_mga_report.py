"""
비용 상한 기반 최소 재생설비 대안(MGA) 분석 보고서(Word) + 그림 + 표 생성

선행: python analyze_mga.py <결과폴더 ...>   (reports/mga_<날짜>/mga_national.csv, mga_regional.csv, mga_summary_all.csv)
사용법:
  python make_mga_report.py
"""
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

D = sorted(glob.glob('reports/mga_*'))[-1]
OUT = D
nat = pd.read_csv(f'{D}/mga_national.csv', index_col=0)
reg = pd.read_csv(f'{D}/mga_regional.csv')
summ = pd.read_csv(f'{D}/mga_summary_all.csv').set_index('case')

REGION_KR = {'BSN': '부산울산', 'CBD': '충북', 'CND': '대전세종충남', 'GBD': '대구경북', 'GGD': '서울경기',
             'GND': '경남', 'GWD': '강원', 'ICN': '인천', 'JBD': '전북', 'JJD': '제주', 'JND': '광주전남'}
REG = ['GGD', 'ICN', 'GWD', 'CBD', 'CND', 'JBD', 'GBD', 'BSN', 'GND', 'JND', 'JJD']
C_PV, C_WT, COL_B0, C_MGA, C_PEN, C_CHP, C_GAS = '#F2B134', '#3C8DBC', '#7F7F7F', '#1F4E79', '#C0392B', '#8E5EA2', '#E07B39'

MGA = [c for c in nat.index if c.startswith('MGA-')]
MGA = sorted(MGA, key=lambda c: float(c.split('-', 1)[1]))
CASES = ['B0'] + MGA
EPS = {c: float(c.split('-', 1)[1]) for c in MGA}
EPSD = {c: (0.01 if e == 0 else e) for c, e in EPS.items()}
PEN = [c for c in ('check_penalty_030', 'check_penalty_100') if c in summ.index]
C_B0 = float(nat.at['B0', 'cost_sys'])
RE_B0 = float(nat.at['B0', 'RE_GW'])
OBJ_B0 = float(summ.at['B0', 'objective_full'])


def V(c, k):
    return float(nat.at[c, k])


def inc(c):
    return 100 * (V(c, 'cost_sys') / C_B0 - 1)


def dre(c):
    return V(c, 'RE_GW') - RE_B0


def label(c):
    return 'B0' if c == 'B0' else f'MGA-{EPS[c]:g}'


figs = {}


def save(fig, name):
    p = os.path.join(OUT, name)
    fig.tight_layout()
    fig.savefig(p, dpi=200)
    plt.close(fig)
    figs[name] = p


# ───────────── 그림 ─────────────
# fig1 비용-설비 경계
fig, ax = plt.subplots(1, 2, figsize=(11, 3.8))
x = [0.0] + [inc(c) for c in MGA]
y = [RE_B0] + [V(c, 'RE_GW') for c in MGA]
ax[0].plot(x, y, '-o', color=C_MGA, label='MGA 최소 설비 경계', zorder=3)
ax[0].scatter([0], [RE_B0], s=70, color=COL_B0, zorder=4, label='B0 (비용 최소)')
for c in MGA:
    off = {0.0: (6, 2), 0.1: (6, 0)}.get(EPS[c], (4, 5))
    ax[0].annotate(f'{dre(c):+.1f}', (inc(c), V(c, 'RE_GW')), textcoords='offset points', xytext=off, fontsize=7.5)
for c, mk in zip(PEN, ('s', 'D')):
    ax[0].scatter(summ.at[c, 'cost_increase_pct'], summ.at[c, 'RE_GW'], marker=mk, s=55, color=C_PEN, zorder=4,
                  label=f"출력제한 페널티 {int(summ.at[c, 'curtail_penalty'])} (점검)")
ax[0].set_xlabel('시스템 비용 증가율 (%, B0 대비)'); ax[0].set_ylabel('PV + 풍력 총 설비 (GW)')
ax[0].set_title('(a) 비용 증가 허용률별 최소 재생설비'); ax[0].grid(alpha=0.3); ax[0].legend(fontsize=7.5)
steps = list(MGA)
prev = ['B0'] + steps[:-1]
SEG = [f"{(EPSD[p] if p != 'B0' else 0):g}→{EPSD[c]:g}%" for p, c in zip(prev, steps)]
mc = [(V(c, 'cost_sys') - V(p, 'cost_sys')) * 1e4 / max(V(p, 'RE_GW') - V(c, 'RE_GW'), 1e-9) for p, c in zip(prev, steps)]
red = [V(p, 'RE_GW') - V(c, 'RE_GW') for p, c in zip(prev, steps)]
xx = np.arange(len(steps))
ax[1].bar(xx, mc, color=C_MGA, width=0.6)
for i, (m_, r_) in enumerate(zip(mc, red)):
    ax[1].text(i, m_ + 20, f'{m_:,.0f}\n(-{r_:.1f} GW)', ha='center', fontsize=7)
ax[1].set_xticks(xx, SEG, fontsize=7.5)
ax[1].set_ylabel('억원/년 per GW 감축'); ax[1].set_ylim(0, max(mc) * 1.25)
ax[1].set_title('(b) 구간별 재생설비 1 GW 감축당 추가비용'); ax[1].grid(axis='y', alpha=0.3)
save(fig, 'fig1_frontier.png')

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
ax.set_xticks(xx, [label(c) for c in CASES]); ax.set_ylabel('GW'); ax.set_ylim(0, 450)
ax2 = ax.twinx()
ax2.plot(xx, [V(c, 'curtail_%') for c in CASES], '-o', color=C_PEN, label='출력제한률')
ax2.set_ylabel('출력제한률 (%)'); ax2.set_ylim(0, 14)
h1, l1 = ax.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
ax.legend(h1 + h2, l1 + l2, fontsize=8, loc='upper right', ncol=3)
ax.set_title('케이스별 PV·풍력 설비와 출력제한률'); ax.grid(axis='y', alpha=0.3)
save(fig, 'fig2_capacity_mix.png')

# fig3 지역별 변화
SHOW = [c for c in ('MGA-1', 'MGA-3', 'MGA-5') if c in MGA]
R = {c: reg[reg.case == c].set_index('region').reindex(REG) for c in CASES}
fig, ax = plt.subplots(1, len(SHOW) + 1, figsize=(13, 3.8), sharey=True)
yy = np.arange(len(REG))[::-1]
for k, c in enumerate(SHOW):
    ax[k].barh(yy + 0.2, R[c].PV_GW - R['B0'].PV_GW, 0.4, color=C_PV, label='PV')
    ax[k].barh(yy - 0.2, R[c].WT_GW - R['B0'].WT_GW, 0.4, color=C_WT, label='풍력')
    ax[k].axvline(0, color='k', lw=0.6); ax[k].grid(axis='x', alpha=0.3)
    ax[k].set_title(f'({chr(97 + k)}) {c} - B0 (GW)'); ax[k].legend(fontsize=7, loc='lower left')
ax[0].set_yticks(yy, [REGION_KR[r] for r in REG])
k = len(SHOW)
for j, c in enumerate(['B0'] + SHOW):
    ax[k].barh(yy + 0.3 - 0.2 * j, R[c]['WT_pot_%'], 0.2, label=c,
               color=[COL_B0, '#9DC3E6', '#2E75B6', '#1F4E79'][j])
ax[k].axvline(100, color=C_PEN, ls='--', lw=1); ax[k].set_xlim(0, 140)
ax[k].set_title(f'({chr(97 + k)}) 풍력 설비 / 기술적 잠재량 (%)'); ax[k].legend(fontsize=7, loc='lower right')
ax[k].grid(axis='x', alpha=0.3)
save(fig, 'fig3_regional.png')

# fig4 메커니즘: B0 대비 전력 수급 변화
items = [('PV 발전', 'PV_TWh', 1), ('풍력 발전', 'WT_TWh', 1), ('출력제한', 'curtail_TWh', 1),
         ('LNG 발전', 'LNG_power_TWh', 1), ('CHP 발전', 'CHP_power_TWh', 1),
         ('히트펌프 전력', 'HP_elec_TWh', 1), ('수전해 전력', 'electrolyser_elec_TWh', 1),
         ('EV 충전(V2G 포함)', 'EV_charge_TWh', 1), ('ESS 충전', 'ESS_charge_TWh', 1)]
fig, ax = plt.subplots(figsize=(11, 3.8))
xx = np.arange(len(items)); wdt = 0.8 / len(SHOW)
for j, c in enumerate(SHOW):
    vals = [V(c, k) - V('B0', k) for _, k, _ in items]
    ax.bar(xx - 0.4 + wdt * (j + 0.5), vals, wdt, label=c, color=['#9DC3E6', '#2E75B6', '#1F4E79'][j])
ax.axhline(0, color='k', lw=0.6)
ax.axvspan(-0.5, 2.5, color=C_PV, alpha=0.08); ax.axvspan(2.5, 4.5, color=C_GAS, alpha=0.08)
ax.axvspan(4.5, 8.5, color=C_WT, alpha=0.08)
ax.text(1, ax.get_ylim()[1] * 0.92, '재생에너지', ha='center', fontsize=8)
ax.text(3.5, ax.get_ylim()[1] * 0.92, '가스(연료 총량 고정)', ha='center', fontsize=8)
ax.text(6.5, ax.get_ylim()[1] * 0.92, '전력 소비(부문결합·저장)', ha='center', fontsize=8)
ax.set_xticks(xx, [n_ for n_, _, _ in items], fontsize=8)
ax.set_ylabel('B0 대비 변화 (TWh)'); ax.legend(fontsize=8, loc='lower left'); ax.grid(axis='y', alpha=0.3)
ax.set_title('재생설비 감축의 메커니즘: B0 대비 연간 전력 수급 변화')
save(fig, 'fig4_mechanism.png')

# fig5 비용 구성 변화
groups = [('PV 투자비', ['c_capex_PV']), ('풍력 투자비', ['c_capex_WT']),
          ('가스 연료비(발전기)', ['c_opex_gas']), ('CHP 운전비', ['c_opex_CHP']),
          ('V2G·ESS·연료전지 운전비', ['c_opex_EV_V2G', 'c_opex_ESS_Ch', 'c_opex_ESS_Disch', 'c_opex_Fuelcell']),
          ('수소설비 투자비', ['c_capex_Electrolyser', 'c_capex_Fuelcell', 'c_capex_H2_Storage'])]
colors = [C_PV, C_WT, C_GAS, C_CHP, '#2E7D32', '#9E9E9E']
SC = [c for c in MGA if EPS[c] >= 0.5]
fig, ax = plt.subplots(figsize=(10.5, 3.8))
xx = np.arange(len(SC))
pos = np.zeros(len(SC)); neg = np.zeros(len(SC))
for (nm, ks), col in zip(groups, colors):
    vals = np.array([sum(nat.at[c, k] - nat.at['B0', k] for k in ks if k in nat and not np.isnan(nat.at[c, k])
                         and not np.isnan(nat.at['B0', k])) for c in SC])
    base = np.where(vals >= 0, pos, neg)
    ax.bar(xx, vals, 0.6, bottom=base, color=col, label=nm)
    pos += np.where(vals >= 0, vals, 0); neg += np.where(vals < 0, vals, 0)
tot = [V(c, 'cost_sys') - C_B0 for c in SC]
ax.plot(xx, tot, 'k_', markersize=26, mew=2.5, label='순 비용 증가')
for i, t in enumerate(tot):
    ax.text(i + 0.33, t, f'+{t:.2f}', va='center', fontsize=8)
ax.axhline(0, color='k', lw=0.6)
ax.set_xticks(xx, SC); ax.set_ylabel('B0 대비 변화 (조원/년)'); ax.grid(axis='y', alpha=0.3)
ax.legend(fontsize=7.5, loc='center left', bbox_to_anchor=(1.01, 0.5)); ax.set_title('비용 구성 변화 (B0 대비)')
save(fig, 'fig5_cost_breakdown.png')

# ───────────── 표 (엑셀) ─────────────
def main_table():
    rows = []
    for c in CASES:
        rows.append({'시나리오': label(c),
                     '비용 증가 허용률': '—' if c == 'B0' else (f'{EPS[c]:g}%' if EPS[c] else '0% (수치허용 0.01%)'),
                     '시스템 비용 (조원)': V(c, 'cost_sys'),
                     '실제 비용 증가율 (%)': 0.0 if c == 'B0' else inc(c),
                     '태양광 GW': V(c, 'PV_GW'), '풍력 GW': V(c, 'WT_GW'),
                     'B0 대비 재생 감소 GW': -dre(c),
                     'ESS GW/GWh': f"{V(c, 'ESS_GW'):.1f} / {V(c, 'ESS_GWh'):.0f}",
                     '출력제한율 (%)': V(c, 'curtail_%')})
    return pd.DataFrame(rows)


def detail_table():
    rows = []
    for c in CASES:
        rows.append({'시나리오': label(c), '비용 상한 (조원)': '—' if c == 'B0' else f"{nat.at[c, 'cost_cap'] / 1e9:.3f}",
                     '1단계 최소 신규 K* (GW)': V(c, 'stage1_K_MW') / 1e3,
                     '최종 PV 신규 (GW)': V(c, 'PV_new_GW'), '최종 풍력 신규 (GW)': V(c, 'WT_new_GW'),
                     '최종 신규 합계 (GW)': V(c, 'RE_new_GW'),
                     '1단계 비용 (조원)': '—' if c == 'B0' else f"{nat.at[c, 'stage1_C_sys'] / 1e9:.3f}",
                     '2단계 비용 (조원)': V(c, 'cost_sys'),
                     '배출량 (Mt)': V(c, 'CO2_Mt'), '전력 slack (TWh)': 0.0,
                     '열보조공급 (TWh)': V(c, 'heat_fallback_TWh')})
    return pd.DataFrame(rows)


def balance_table(cols):
    items = [('PV 발전량', 'PV_TWh'), ('풍력 발전량', 'WT_TWh'), ('출력제한량', 'curtail_TWh'),
             ('PV 이용률 (%)', 'PV_CF_%'), ('원자력', 'nuclear_TWh'), ('LNG 발전', 'LNG_power_TWh'),
             ('CHP 발전', 'CHP_power_TWh'), ('CHP 열', 'CHP_heat_TWh'), ('CHP 연료 (가스)', 'CHP_fuel_TWh'),
             ('가스 연료 총량', 'gas_fuel_TWh'), ('히트펌프 전력', 'HP_elec_TWh'), ('히트펌프 열', 'HP_heat_TWh'),
             ('수전해 전력', 'electrolyser_elec_TWh'), ('연료전지 발전', 'fuelcell_power_TWh'),
             ('EV 충전 (V2G 포함)', 'EV_charge_TWh'), ('V2G 방전', 'V2G_TWh'), ('ESS 충전', 'ESS_charge_TWh'),
             ('전력 소비 합계', 'elec_demand_total_TWh'), ('HVDC 수송량', 'HVDC_TWh'),
             ('CO2 잠재가격 (천원/t)', 'CO2_price'), ('부하가중 전력가격 (천원/MWh)', 'EL_price_loadw')]
    df = pd.DataFrame({'항목 (TWh)': [a for a, _ in items]})
    for c in cols:
        df[label(c)] = [V(c, k) for _, k in items]
    return df


def regional_table(cols):
    rows = []
    for r in REG:
        row = {'지역': REGION_KR[r]}
        for c in cols:
            row[f'{label(c)} PV'] = R[c].at[r, 'PV_GW']
            row[f'{label(c)} 풍력'] = R[c].at[r, 'WT_GW']
        row['풍력 잠재량'] = TECH_POTENTIAL_GW['WT'][r]
        rows.append(row)
    return pd.DataFrame(rows)


def penalty_table():
    rows = []
    for c in PEN:
        s = summ.loc[c]
        ci = float(s['cost_increase_pct'])
        # 같은 비용에서 MGA 경계의 재생 감소량(선형 보간)
        xs = [0.0] + [inc(m) for m in MGA]; ys = [0.0] + [-dre(m) for m in MGA]
        rows.append({'점검 케이스': f"출력제한 페널티 {int(s['curtail_penalty'])}",
                     '실제 비용 증가율 (%)': ci,
                     '참고: 열보조 페널티 포함 기준 (%)': 100 * (float(s['objective_full']) / OBJ_B0 - 1),
                     'PV / 풍력 (GW)': f"{s['PV_GW']:.1f} / {s['WT_GW']:.1f}",
                     '재생 감소 (GW)': RE_B0 - float(s['RE_GW']),
                     '같은 비용의 MGA 감소 (GW)': float(np.interp(ci, xs, ys)),
                     '출력제한율 (%)': float(s['curtail_%'])})
    return pd.DataFrame(rows)


TBL = {'표1_결과요약': main_table(), '표2_단계별상세': detail_table(), '표3_페널티점검': penalty_table(),
       '표4_수급변화': balance_table(CASES), '표5_지역별설비': regional_table(['B0'] + SHOW)}
with pd.ExcelWriter(os.path.join(OUT, 'MGA_분석_표.xlsx')) as xw:
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


def g(c, k):
    return V(c, k) - V('B0', k)


m1, m3, m5 = 'MGA-1', 'MGA-3', 'MGA-5'
m05 = 'MGA-0.5'
m0 = 'MGA-0' if 'MGA-0' in MGA else None
m01 = 'MGA-0.1' if 'MGA-0.1' in MGA else None
pen = penalty_table().set_index('점검 케이스')
mc_tbl = pd.DataFrame({'구간': SEG,
                       '추가비용 (조원/년)': [V(c, 'cost_sys') - V(p, 'cost_sys') for p, c in zip(prev, steps)],
                       '재생 감소 (GW)': red, '1 GW 감축당 비용 (억원/년)': mc})

# 표지·개요
doc.add_heading('비용 증가를 제한한 최소 재생설비 대안 분석 (MGA)', 0)
P('2040년 재생에너지 보급 상한 해제 시나리오 기준 — 비용 증가 허용률 0~5%에서 태양광·풍력 신규 설비 최소화', size=12)
P(f'작성일: {datetime.now():%Y-%m-%d}   |   모형: PyPSA 0.31.1 + linopy + CPLEX Barrier, 8,784시간 단일연도', size=9)

H('개요', 1)
P('기존 비용 최적해(B0)는 태양광을 많이 설치하는 조합을 선택한다. 이 분석은 시스템 비용을 B0보다 조금 더 허용할 때 '
  '태양광·풍력 신규 설비를 얼마나 줄일 수 있는지, 그리고 그 대안이 어떤 설비·운영 구성인지를 확인한다. '
  'Modelling to Generate Alternatives(MGA) 방식으로, 비용 상한 C ≤ (1+ε)·C_B0 아래에서 PV·풍력 신규 용량을 최소화한 뒤, '
  '같은 용량 수준에서 비용을 다시 최소화하였다.')
T(pd.DataFrame([
    ['기준 입력', '현재 interface (2040년, 재생 보급 상한 해제). 입력 수치는 변경하지 않음'],
    ['B0', '비용 최소화. 전력 slack 0 고정 (원래 결과에서도 slack 0 → 결과 동일)'],
    ['MGA-ε', f"ε = {', '.join(f'{EPS[c]:g}' for c in MGA)}% — 모든 케이스가 B0 비용을 기준으로 별도 실행"],
    ['대체 목적함수', 'Σ (태양광 + 풍력) 신규 용량 [GW] 최소화 (11개 지역 × 2기술 = 22개 확장 설비)'],
    ['2단계', '신규 용량 ≤ K*(1+10⁻⁴) 제약 하에서 원래 비용 최소화 → 최소 설비 대안 중 가장 저렴한 구성'],
    ['비용 정의', 'PyPSA 목적함수(신규 설비 연간화 투자비 + 전체 운전비) − 열보조공급 페널티. B0·상한·결과표 모두 동일'],
    ['점검 케이스', '출력제한 페널티 30·100 천원/MWh (현재 입력으로 재계산)'],
], columns=['항목', '내용']), font=8, caption='표 0-1. 분석 구성')

H('핵심 결과', 2)
mc01 = (V(m01, 'cost_sys') - C_B0) * 1e4 / -dre(m01) if m01 else float('nan')
B([('1. 비용 최적해 근처는 평평하다. ',
    (f"동일 비용(MGA-0, 수치 허용 0.01%)에서도 {-dre(m0):.1f} GW, 비용 +0.1%(약 {V(m01, 'cost_sys') - C_B0:.2f}조원/년)로 "
     f"{-dre(m01):.1f} GW를 줄일 수 있다. 비용 최적해는 태양광 쪽으로 기울어 있지만, 거의 같은 비용의 대안이 존재한다.")
    if m0 and m01 else ''),
   ('2. 비용을 0.5~1% 더 허용하면 재생설비가 크게 준다. ',
    f"MGA-0.5에서 PV·풍력 총설비가 {RE_B0:.1f} → {V(m05, 'RE_GW'):.1f} GW로 {-dre(m05):.1f} GW 감소한다 "
    f"(추가비용 {V(m05, 'cost_sys') - C_B0:.2f}조원/년). 1%는 {-dre(m1):.1f} GW, 3%는 {-dre(m3):.1f} GW, "
    f"5%는 {-dre(m5):.1f} GW 감소한다. 감축 효과는 빠르게 체감한다. 1 GW 감축당 추가비용이 +0.1%까지는 평균 {mc01:,.0f}억원/년이지만 "
    f"4→5% 구간은 {mc[-1]:,.0f}억원/년이며, 3% 이후에는 GW당 1천억원/년을 넘는다."),
   ('3. 대안의 핵심은 태양광→풍력 전환이다. 신규 ESS는 늘지 않는다. ',
    f"MGA-5에서 PV {g(m5, 'PV_GW'):+.1f} GW, 풍력 {g(m5, 'WT_GW'):+.1f} GW. 풍력 이용률(약 23.5%)이 PV(약 13%)의 "
    '1.8배라 같은 무탄소 발전량을 적은 GW로 얻는다. 어느 케이스에서도 신규 ESS는 0이고, 기존 ESS 충전량도 오히려 줄어든다. '
    'PV 잉여가 줄어 저장 수요 자체가 감소하기 때문이다.'),
   ('4. 잉여전력 순환이 줄어 필요한 재생 발전량 자체가 감소한다. ',
    f"MGA-5에서 출력제한 {V('B0', 'curtail_TWh'):.0f} → {V(m5, 'curtail_TWh'):.0f} TWh, "
    f"수전해 전력 {g(m5, 'electrolyser_elec_TWh'):+.1f}, V2G 방전 {g(m5, 'V2G_TWh'):+.1f} TWh로 "
    f"전력 소비가 {-g(m5, 'elec_demand_total_TWh'):.1f} TWh 줄고, 재생 발전량도 {-(g(m5, 'PV_TWh') + g(m5, 'WT_TWh')):.1f} TWh 덜 필요하다."),
   ('5. 3% 이상에서는 가스 사용처가 LNG 발전에서 CHP로 이동한다. ',
    f"CO2 상한(58.87 Mt)과 가스 연료 총량({V('B0', 'gas_fuel_TWh'):.0f} TWh)은 그대로인 채, CHP 연료가 "
    f"{V('B0', 'CHP_fuel_TWh'):.1f} → {V(m3, 'CHP_fuel_TWh'):.1f} TWh로 늘고 CHP 열이 히트펌프 열을 대체한다 "
    f"(히트펌프 전력 {g(m3, 'HP_elec_TWh'):+.1f} TWh). 같은 탄소로 전력 수요를 줄이는 우회 경로이다."),
   ('6. 주의: 늘어나는 풍력이 서울경기에 집중되어 기술적 잠재량을 넘는다. ',
    f"서울경기 풍력이 B0 {R['B0'].at['GGD', 'WT_GW']:.1f} → MGA-1 {R[m1].at['GGD', 'WT_GW']:.1f}, "
    f"MGA-5 {R[m5].at['GGD', 'WT_GW']:.1f} GW(잠재량 {TECH_POTENTIAL_GW['WT']['GGD']:.1f} GW의 "
    f"{R[m5].at['GGD', 'WT_pot_%']:.0f}%)가 된다. 상한 해제 시나리오의 특성이므로, 정책 수치로 쓰려면 "
    '잠재량 상한을 둔 조건에서 같은 분석을 반복할 필요가 있다.'),
   ('7. 구현 점검 통과. ',
    f"현재 입력으로 다시 계산한 출력제한 페널티 30(비용 +{pen.iloc[0]['실제 비용 증가율 (%)']:.2f}%, 재생 "
    f"−{pen.iloc[0]['재생 감소 (GW)']:.1f} GW)과 100(+{pen.iloc[1]['실제 비용 증가율 (%)']:.2f}%, "
    f"−{pen.iloc[1]['재생 감소 (GW)']:.1f} GW)은 모두 MGA 경계보다 위에 있다. 즉 같은 비용에서 MGA가 더 많이 줄인다"
    f"(각각 −{pen.iloc[0]['같은 비용의 MGA 감소 (GW)']:.1f}, −{pen.iloc[1]['같은 비용의 MGA 감소 (GW)']:.1f} GW).")])

T(main_table(), font=8, caption='표 0-2. 결과 요약 (요청 양식)')
F('fig1_frontier.png', '그림 1. 비용 증가 허용률별 최소 재생설비 경계(a)와 구간별 1 GW 감축당 추가비용(b)')

# 1장 방법
doc.add_page_break()
H('1. 분석 방법', 1)
H('1.1 단계별 계산', 2)
B([('B0: ', f"현재 interface 입력 그대로 비용을 최소화한다. 시스템 비용 C_B0 = {C_B0:.4f}조원/년 "
    f"(PyPSA 목적함수 {OBJ_B0 / 1e9:.4f}조원 − 열보조공급 페널티 {OBJ_B0 / 1e9 - C_B0:.4f}조원). "
    'PyPSA statistics(capex+opex−기존설비 고정비)로 따로 계산한 값과 일치함을 확인하였다.'),
   ('MGA 1단계: ', '원래 비용식을 보관한 뒤 C ≤ (1+ε)·C_B0 제약을 추가하고, 목적함수를 태양광·풍력 신규 용량 합계 최소화로 교체한다. '
    '결과 K*는 그 비용 한도에서 가능한 최소 신규 용량이다.'),
   ('MGA 2단계: ', '신규 용량 ≤ K*(1+10⁻⁴) 제약을 두고 목적함수를 원래 비용으로 되돌려 다시 최소화한다. '
    '1단계의 목적값은 GW 단위의 설비용량이므로 비용으로 읽지 않고, 보고하는 비용은 2단계 값이다.'),
   ('기준 고정: ', '모든 케이스는 같은 B0 입력에서 따로 실행하며, 상한은 항상 최초 B0 비용 기준이다(직전 케이스 비용을 기준으로 쓰지 않음).')])
H('1.2 공통 제약과 구현', 2)
B([('기존 사용자 정의 제약 보존: ', 'PyPSA_GUI.optimize_network를 그대로 호출하고, 그 안의 extra_functionality(CHP 포함 CO2 통합제약, '
    'DR 연간 발동시간, DC/Fab RE 조달 제약) 뒤에 MGA 제약과 목적함수 교체만 덧붙였다. PyPSA의 optimize_mga는 '
    'extra_functionality를 받지 않아 사용자 정의 제약이 빠지므로 사용하지 않았다. PyPSA_GUI.py와 입력 파일은 수정하지 않았다.'),
   ('전력 slack: ', '전력 slack 22개(지역 EL slack 11, EV slack 11)의 발전을 0으로 고정. B0 원래 결과도 slack 0이어서 B0 비용은 '
    '변하지 않았다(61.7234조원, 기존 수요 민감도 기준 케이스와 동일).'),
   ('열보조공급(H_Fallback): ', f"B0에서 {V('B0', 'heat_fallback_TWh'):.2f} TWh가 쓰이며 페널티 단가(5,000천원/MWh)가 목적함수에 "
    f"{OBJ_B0 / 1e9 - C_B0:.2f}조원 들어 있다. 이 금액은 실제 시스템 비용에서 제외하고, MGA에서는 지역별 열보조공급량을 B0 이하로 제한하였다. "
    '그렇지 않으면 추가 예산으로 열보조공급을 늘려 히트펌프 전력수요를 줄이는 우회가 가능해진다. 결과적으로 전 케이스에서 B0와 같은 양이 쓰였다.'),
   ('신규 설비 정의: ', '분석 시작 시점 기존 설비(입력 p_nom: PV 134.0 GW, 풍력 37.0 GW)를 제외한 추가 설치량. B0 최적 설비를 하한으로 고정하지 않았다.'),
   ('CO2: ', '배출 상한 58.87 Mt는 모든 케이스에서 구속 조건이며 배출량은 상한과 같다. CHP 연료는 LNG 연료공급 발전기를 통해 '
    '상한 집계에 포함된다.')])
H('1.3 수치 처리', 2)
B([('MGA-0: ', '비용 상한이 최소비용과 같으면 실행 가능 영역이 거의 한 점이 되어 내부점법이 수렴하지 않았다'
    '(상대 허용오차 10⁻⁵에서 수치 오류로 종료). 그래서 MGA-0은 상대 허용오차 10⁻⁴(0.01%, 약 54억원/년)를 두고 계산하였다. '
    '이 경우에도 1단계는 Barrier 수렴 기준 10⁻⁶에서 끝나지 않아 10⁻⁵로 재시도하여 풀었다(이전 분석에서 두 기준의 결과가 같음을 확인).'
    + (f" 추가로 0과 0.5% 사이의 형태를 보기 위해 MGA-0.1을 계산하였다." if m01 else '')),
   ('해 반영 확인: ', "솔버가 'unknown' 상태로 끝나면 해가 네트워크에 반영되지 않으므로, 해 반영 여부를 검사하고 실패하면 "
    '수렴 허용오차를 10⁻⁵로 완화해 재시도하도록 하였다. 최종 결과는 모두 정상(optimal) 종료된 해이다.'),
   ('비용 상한 식 스케일링: ', '비용 상한 제약을 10⁶으로 나누어(십억원 단위) 수치 조건을 개선하였다.')])
T(detail_table(), font=7.5, caption='표 1-1. 단계별 상세 결과')

# 2장 결과
doc.add_page_break()
H('2. 비용 허용률별 최소 재생설비', 1)
F('fig2_capacity_mix.png', '그림 2. 케이스별 PV·풍력 설비와 출력제한률')
B([('MGA-0 (동일 비용)·MGA-0.1: ', (f"수치 허용오차 0.01%(약 {V(m0, 'cost_sys') - C_B0:.3f}조원/년) 안에서 PV·풍력이 "
    f"{-dre(m0):.1f} GW, 0.1%에서 {-dre(m01):.1f} GW 줄어든다. 감소분은 PV(각각 {g(m0, 'PV_GW'):+.1f}, {g(m01, 'PV_GW'):+.1f} GW)이고 "
    f"풍력이 일부({g(m0, 'WT_GW'):+.1f}, {g(m01, 'WT_GW'):+.1f} GW) 대신한다. 최적해 부근의 비용 곡면이 평평해, "
    '비용 최적해의 PV 규모는 비용 가정의 작은 차이에도 크게 달라질 수 있는 값임을 뜻한다.') if m0 and m01 else '수치 문제로 산정하지 못함'),
   ('0.5~1%: ', f"가장 효율적인 구간이다. 0.5%에 {-dre(m05):.1f} GW, 1%에 {-dre(m1):.1f} GW가 줄고, "
    f"감소분은 대부분 광주전남·대전세종충남·부산울산 PV이다."),
   ('2~5%: ', f"추가 감소는 1%p당 약 4~8 GW로 둔화된다. 5%에서도 총설비는 {V(m5, 'RE_GW'):.0f} GW로, "
    f"B0 대비 {100 * -dre(m5) / RE_B0:.0f}% 감소에 그친다."),
   ('출력제한: ', f"{V('B0', 'curtail_%'):.1f}% → MGA-1 {V(m1, 'curtail_%'):.1f}% → MGA-5 {V(m5, 'curtail_%'):.1f}%. "
    '줄어드는 출력제한은 대부분 PV 출력제한이다.')])
T(mc_tbl, font=8, caption='표 2-1. 구간별 재생설비 감축 효율')

# 3장 메커니즘
doc.add_page_break()
H('3. 감축 메커니즘', 1)
F('fig4_mechanism.png', '그림 3. B0 대비 연간 전력 수급 변화 (MGA-1·3·5)')
B([('태양광→풍력: ', '풍력 연간화 투자비(336,746천원/MW)는 PV(73,932천원/MW)의 4.6배이지만 이용률은 1.8배이다. 비용 최소화에서는 PV가 '
    '선택되지만, GW를 줄이는 것이 목표가 되면 풍력이 유리해진다. 비용 증가분은 대부분 이 대체 비용이다.'),
   ('잉여전력 순환 감소: ', 'B0에서는 낮 시간 PV 잉여를 수전해·V2G·ESS로 돌려 쓰고 일부는 버린다. PV가 줄면 이 순환 손실과 출력제한이 함께 줄어 '
    '필요한 재생 발전량 자체가 감소한다(MGA-5에서 재생 발전량 약 9 TWh 감소).'),
   ('ESS: ', '신규 ESS는 모든 케이스에서 0이다. ESS는 PV 잉여를 옮기는 수단인데, 대안 구성에서는 PV 잉여가 줄어 저장 가치가 오히려 낮아진다. '
    'GW 최소화 목표에서도 ESS(저장용량 연간 116,712천원/MWh)는 풍력 전환보다 비싼 수단이다.'),
   ('가스의 CHP 전환 (3% 이상): ', 'CO2 상한 때문에 가스 연료 총량은 고정된다. 같은 가스를 LNG 발전(전기 55.9%) 대신 CHP(전기 약 49% + 열 약 40%)에 쓰면, '
    'CHP 열이 히트펌프 열을 대체해(COP 3) 전력 수요가 준다. 연료 1 MWh당 전력 환산 약 0.62 MWh로 LNG 발전보다 커서, '
    '탄소 한도 안에서 재생 필요량을 줄이는 경로가 된다. 서울경기에서 CHP 열 증가가 가장 크다. 다만 CHP 운전비(연료 포함, '
    '연료 MWh당 80천원)가 LNG 발전 연료비(연료 MWh당 약 55천원)보다 비싸서 그림 4의 CHP 운전비 증가로 나타난다.'),
   ('가격 지표 해석 주의: ', 'MGA 2단계에서는 신규 재생 용량 상한이 구속되므로 CO2 잠재가격과 전력가격(표 3-1)이 B0보다 크게 오른다. '
    '이는 재생 확대를 인위적으로 막은 데 따른 잠재가격이며, 시장 가격 전망으로 해석하면 안 된다.')])
T(balance_table(['B0'] + [c for c in MGA if EPS[c] >= 0.5]), font=7.5, caption='표 3-1. 전력·열·가스 수급 변화 (TWh)', fmt='{:,.1f}')
F('fig5_cost_breakdown.png', '그림 4. B0 대비 비용 구성 변화')

# 4장 지역
doc.add_page_break()
H('4. 지역별 설비 변화', 1)
F('fig3_regional.png', '그림 5. 지역별 PV·풍력 변화와 풍력 잠재량 대비 설비 비율')
B([('PV 감소: ', f"광주전남 {R['B0'].at['JND', 'PV_GW']:.1f} → {R[m5].at['JND', 'PV_GW']:.1f} GW, "
    f"대전세종충남 {R['B0'].at['CND', 'PV_GW']:.1f} → {R[m5].at['CND', 'PV_GW']:.1f} GW, "
    f"부산울산 {R['B0'].at['BSN', 'PV_GW']:.1f} → {R[m5].at['BSN', 'PV_GW']:.1f} GW (MGA-5). "
    '수도권과 멀고 송전 혼잡이 큰 지역의 PV부터 줄어든다(HVDC 수송량도 감소).'),
   ('풍력 증가: ', f"거의 전부 서울경기(B0 {R['B0'].at['GGD', 'WT_GW']:.1f} → MGA-5 {R[m5].at['GGD', 'WT_GW']:.1f} GW)이다. "
    '풍력 비용이 지역 공통이라 수요지 인근이 유리하기 때문이다.'),
   ('잠재량 초과: ', f"서울경기 풍력은 MGA-1부터 기술적 잠재량({TECH_POTENTIAL_GW['WT']['GGD']:.1f} GW)을 넘는다"
    f"(MGA-1 {R[m1].at['GGD', 'WT_pot_%']:.0f}%, MGA-5 {R[m5].at['GGD', 'WT_pot_%']:.0f}%). 인천 PV도 B0부터 잠재량의 "
    f"{R['B0'].at['ICN', 'PV_pot_%']:.0f}%이다. 상한 해제 시나리오의 한계이며, 실제 입지 가능한 대안 규모는 이보다 작을 수 있다.")])
T(regional_table(['B0'] + SHOW), font=7, caption='표 4-1. 지역별 PV·풍력 설비 (GW)', fmt='{:,.1f}')

# 5장 점검
doc.add_page_break()
H('5. 구현 점검: 출력제한 페널티 결과와 비교', 1)
P('출력제한 페널티 분석의 결과(같은 물리 제약, 같은 비용 집계)는 MGA 경계보다 위에 있어야 한다. 즉 같은 비용에서 MGA가 '
  '최소한 그만큼은 재생설비를 줄여야 한다. 이전 페널티 분석은 interface 갱신 전 입력(일반 전력수요 562.6 TWh, 반도체 169.0 TWh)으로 '
  '계산되었으므로, 현재 입력(572.6 TWh, 171.0 TWh)으로 페널티 30·100을 다시 계산해 비교하였다.')
T(penalty_table(), font=7.5, caption='표 5-1. 출력제한 페널티 점검 케이스와 MGA 비교')
B([('비용 기준 차이: ', '이전 보고의 +0.44%·+1.76%는 열보조공급 페널티가 포함된 목적함수 기준이다. 같은 기준으로 다시 계산하면 '
    f"+{pen.iloc[0]['참고: 열보조 페널티 포함 기준 (%)']:.2f}%·+{pen.iloc[1]['참고: 열보조 페널티 포함 기준 (%)']:.2f}%로 이전과 비슷하다. "
    '이 보고서는 페널티를 뺀 실제 시스템 비용을 기준으로 하므로 증가율이 조금 더 크게 나온다.'),
   ('판정: ', f"페널티 30은 MGA-0.5 한도(0.5%) 안에 있고, 감소량({pen.iloc[0]['재생 감소 (GW)']:.1f} GW)이 MGA-0.5 결과"
    f"({-dre(m05):.1f} GW)보다 작다. 페널티 100은 비용 +{pen.iloc[1]['실제 비용 증가율 (%)']:.2f}%로 MGA-2 한도를 약간 넘지만, "
    f"MGA-2가 더 적은 비용으로 더 많이({-dre('MGA-2'):.1f} GW) 줄인다. 구현이 의도대로 동작함을 확인하였다."),
   ('페널티 방식과 차이: ', '페널티는 출력제한 자체를 비싸게 만들어 ESS 충전(약 55 TWh)을 늘리는 반면, MGA 대안은 ESS 이용을 줄이고 '
    '풍력 전환과 CHP 활용으로 설비를 줄인다.')])

# 6장 시사점
doc.add_page_break()
H('6. 종합 및 시사점', 1)
B([('정책 메시지: ', f"재생설비 보급 부담을 줄이는 데 큰 비용이 들지 않는 구간이 있다. 시스템 비용 +0.5%(약 {V(m05, 'cost_sys') - C_B0:.2f}조원/년)로 "
    f"{-dre(m05):.0f} GW, +1%(약 {V(m1, 'cost_sys') - C_B0:.2f}조원/년)로 {-dre(m1):.0f} GW를 줄일 수 있다. 그 이상은 효과가 빠르게 둔화된다."),
   ('수단: ', '설비 감축의 주 수단은 풍력 비중 확대이다. ESS 확대는 설비 감축 수단이 아니다(ESS 비용 민감도·페널티 분석 결과와 일치). '
    '비용을 더 허용하면 열부문(CHP·히트펌프)과 수소·V2G 운영이 함께 조정된다.'),
   ('보급 목표 설정: ', 'PV 보급량을 낮추는 대안은 풍력 보급 목표(특히 수도권 인근 육상·해상 풍력)와 함께 제시되어야 실현 가능하다.'),
   ('한계: ', '① 상한 해제 시나리오라 서울경기 풍력이 잠재량을 넘는다. ② 대체 목적함수가 PV와 풍력 GW를 같은 가중치로 본다'
    '(면적·입지 부담 차이는 반영하지 않음). ③ MGA 대안의 가격 지표는 용량 제한에 따른 잠재가격이다.')])
H('추가로 해 볼 만한 분석', 2)
B([('1. 잠재량 상한(20%) 조건 MGA: ', '서울경기 풍력 집중이 막혔을 때 감축 가능량과 비용 곡선이 어떻게 바뀌는지 확인'),
   ('2. 가중 목적함수: ', 'PV와 풍력의 면적·입지 부담을 반영한 가중치(예: GW당 필요 면적)로 최소화'),
   ('3. 풍력·PV 별도 최소화: ', 'PV 신규만 최소화(풍력 자유)하여 PV 보급 부담의 최소치를 산정'),
   ('4. 열부문 제약 민감도: ', 'CHP 확대를 제한했을 때 3% 이상 구간의 감축량 변화 확인')])

# 부록
doc.add_page_break()
H('부록. 산출 파일', 1)
B(['결과 폴더: results/mga_sweep_* (B0.nc, MGA-*.nc, mga_summary.csv)',
   f'분석 표: {OUT}/MGA_분석_표.xlsx (요약·단계별·점검·수급·지역 + 전체 지표)',
   '실행 스크립트: run_mga_sweep.py (MGA 실행), analyze_mga.py (집계), make_mga_report.py (보고서)'])

path = os.path.join(OUT, 'MGA_최소설비대안_분석_보고서.docx')
doc.save(path)
print(f'완료: {path}')
print(main_table().round(3).to_string(index=False))
print(penalty_table().round(3).to_string(index=False))
print(mc_tbl.round(1).to_string(index=False))
