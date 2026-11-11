"""
수요 감축 민감도 분석 보고서(Word) + 그림(PNG) + 표(Excel) 생성

사용법:
  python make_demand_report.py
출력:
  reports/demand_sensitivity_<날짜>/수요감축_민감도분석_보고서.docx, 수요감축_민감도분석_표.xlsx, fig*.png
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
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from run_demand_sweep import target_loads

logging.disable(logging.WARNING)
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 9

R = 'results'
DIRS = [rf"{R}\demand_sweep_20261029_164849", rf"{R}\demand_sweep_20261102_101935"]
BASE_NC = rf"{DIRS[0]}\base\base.nc"
OUT = os.path.join('reports', f"demand_sensitivity_{datetime.now().strftime('%Y%m%d')}")
os.makedirs(OUT, exist_ok=True)

REGION_KR = {'BSN': '부산울산', 'CBD': '충북', 'CND': '대전세종충남', 'GBD': '대구경북', 'GGD': '서울경기',
             'GND': '경남', 'GWD': '강원', 'ICN': '인천', 'JBD': '전북', 'JJD': '제주', 'JND': '광주전남'}
C_PV, C_WT, C_LINE, C_CO2 = '#F2B134', '#3C8DBC', '#C0392B', '#2E7D32'
C_EL, C_DC, C_FAB = '#5B7DB1', '#E07B39', '#8E5EA2'

CASE_KR = {'base': '기준', 'EL_m05': '일반 −5%', 'EL_m10': '일반 −10%', 'EL_m15': '일반 −15%',
           'DC_m25': 'DC −25%', 'DC_m50': 'DC −50%', 'Fab_m25': 'Fab −25%', 'Fab_m50': 'Fab −50%',
           'DCFab_m25': 'DC+Fab −25%', 'DCFab_m50': 'DC+Fab −50%'}
OFF = {'Fab_m25': (-48, 6), 'Fab_m50': (-48, 6), 'DCFab_m25': (4, -11)}
SET1 = ['base', 'EL_m05', 'EL_m10', 'EL_m15']
SET2 = ['base', 'DC_m25', 'DC_m50', 'Fab_m25', 'Fab_m50', 'DCFab_m25', 'DCFab_m50']


# ───────────────────────────── 데이터 ─────────────────────────────
def load_both(fname):
    a, b = (pd.read_csv(os.path.join(d, fname)) for d in DIRS)
    return pd.concat([a[~a.case.isin(b.case)], b], ignore_index=True)


s = load_both('sweep_summary.csv').set_index('case').loc[list(CASE_KR)]
s['demand_TWh'] = s[['demand_EL_TWh', 'demand_DC_TWh', 'demand_Fab_TWh']].sum(axis=1)
s['cost'] = s.objective / 1e9
s['RE_GW'] = s.PV_GW + s.WT_GW
b = s.loc['base']
s['d_demand'] = b.demand_TWh - s.demand_TWh
s['d_cost'] = b.cost - s.cost
s['avoided'] = (s.d_cost * 1e9 / (s.d_demand * 1e6)).where(s.d_demand > 0)
s['d_REgen'] = b.RE_gen_TWh - s.RE_gen_TWh
s['label'] = [CASE_KR[c] for c in s.index]
s['plabel'] = s.label.str.replace('−', '-')

mix = load_both('mix_by_carrier.csv')
reg = load_both('regional_capacity.csv')
reg['k'] = reg.tech + ':' + reg.region
opt = reg.pivot_table(index='k', columns='case', values='opt_GW')[list(CASE_KR)]
exist = reg.drop_duplicates('k').set_index('k').existing_GW

nb = pypsa.Network(BASE_NC)
w = nb.snapshot_weightings.generators
P = nb.get_switchable_as_dense('Load', 'p_set')
dem_reg = []
for kind in ('EL', 'DC', 'Fab'):
    for l in target_loads(nb, (kind,)):
        dem_reg.append({'kind': kind, 'region': l.strip()[:3], 'TWh': float((P[l] * w).sum()) / 1e6})
dem_reg = pd.DataFrame(dem_reg).pivot_table(index='region', columns='kind', values='TWh')[['EL', 'DC', 'Fab']]
prof = {k: P[target_loads(nb, (k,))].sum(axis=1) for k in ('EL', 'DC', 'Fab')}
lf = {k: float(v.mean() / v.max()) for k, v in prof.items()}
daily = {k: v.groupby(v.index.hour).mean() / v.mean() for k, v in prof.items()}
monthly = {k: v.groupby(v.index.month).mean() / v.mean() for k, v in prof.items()}
co2_cap = float(nb.global_constraints.at['CO2_limit', 'constant']) / 1e6

excel_tables = {}
figs = {}


def save(fig, name):
    p = os.path.join(OUT, name)
    fig.tight_layout()
    fig.savefig(p, dpi=200)
    plt.close(fig)
    figs[name] = p


def summary_table(cases):
    d = s.loc[cases]
    return pd.DataFrame({
        '케이스': d.label.values, '총수요 (TWh)': d.demand_TWh.values, '수요 감소 (TWh)': d.d_demand.values,
        'PV (GW)': d.PV_GW.values, '풍력 (GW)': d.WT_GW.values, '재생 발전량 (TWh)': d.RE_gen_TWh.values,
        '출력제한률 (%)': d['curtail_%'].values, '가스 발전 (TWh)': d.gas_gen_TWh.values,
        'CO2 그림자가격 (천원/t)': d.CO2_shadow_price.values, '평균 전력가격 (천원/MWh)': d.EL_price_mean.values,
        '총비용 (조원)': d.cost.values, '비용 감소 (조원)': d.d_cost.values,
        'MWh당 회피비용 (천원)': d.avoided.values})


FMT = {'총비용 (조원)': '{:.2f}', '비용 감소 (조원)': '{:.2f}', 'MWh당 회피비용 (천원)': '{:.0f}',
       'CO2 그림자가격 (천원/t)': '{:.0f}', '평균 전력가격 (천원/MWh)': '{:.1f}'}

# ───────────────────────────── 그림 ─────────────────────────────
# 그림 0: 수요 구성
fig, ax = plt.subplots(1, 3, figsize=(11, 3.4), gridspec_kw={'width_ratios': [1.5, 1, 1]})
rr = dem_reg.sort_values('EL', ascending=False)
xx = np.arange(len(rr))
bot = np.zeros(len(rr))
for k, c, lb in (('EL', C_EL, '일반'), ('DC', C_DC, '데이터센터'), ('Fab', C_FAB, '팹')):
    ax[0].bar(xx, rr[k], bottom=bot, color=c, label=lb)
    bot += rr[k].values
ax[0].set_xticks(xx, [REGION_KR[r] for r in rr.index], rotation=45, ha='right', fontsize=7.5)
ax[0].set_ylabel('연간 수요 (TWh)'); ax[0].set_title('(a) 지역별 전력수요 구성'); ax[0].legend(fontsize=7.5)
for k, c, lb in (('EL', C_EL, '일반'), ('DC', C_DC, '데이터센터'), ('Fab', C_FAB, '팹')):
    ax[1].plot(daily[k].index, daily[k].values, color=c, label=f'{lb} (부하율 {lf[k]:.2f})')
    ax[2].plot(monthly[k].index, monthly[k].values, 'o-', color=c, ms=3, label=lb)
ax[1].set_xlabel('시각'); ax[1].set_ylabel('평균 대비 비율'); ax[1].set_title('(b) 시간대별 평균 부하 형태')
ax[1].legend(fontsize=7); ax[1].grid(alpha=0.3)
ax[2].set_xlabel('월'); ax[2].set_title('(c) 월별 평균 부하 형태'); ax[2].grid(alpha=0.3); ax[2].set_xticks(range(1, 13))
save(fig, 'fig0_demand_profile.png')


def mix_fig(cases, name, title_x):
    d = s.loc[cases]
    fig, ax = plt.subplots(1, 2, figsize=(10, 3.6))
    x = np.arange(len(d))
    ax[0].bar(x, d.PV_GW, color=C_PV, label='PV')
    ax[0].bar(x, d.WT_GW, bottom=d.PV_GW, color=C_WT, label='풍력')
    for i, v in enumerate(d.RE_GW):
        ax[0].text(i, v + 5, f'{v:.0f}', ha='center', fontsize=8)
    ax[0].set_xticks(x, d.plabel, rotation=30 if len(d) > 4 else 0, ha='right' if len(d) > 4 else 'center', fontsize=8)
    ax[0].set_ylabel('설비용량 (GW)'); ax[0].set_ylim(0, 450); ax[0].set_title('(a) 재생에너지 설비용량')
    ax[0].legend(loc='lower left', fontsize=8)
    ax[1].bar(x, d.cost, color='#8C9BAB', label='총비용 (조원)')
    for i, v in enumerate(d.cost):
        ax[1].text(i, v / 2, f'{v:.1f}', ha='center', fontsize=8, color='white', fontweight='bold')
    ax[1].set_ylim(0, 75); ax[1].set_ylabel('총비용 (조원/년)')
    ax[1].set_xticks(x, d.plabel, rotation=30 if len(d) > 4 else 0, ha='right' if len(d) > 4 else 'center', fontsize=8)
    a2 = ax[1].twinx()
    a2.plot(x, d.CO2_shadow_price, 'o-', color=C_CO2, label='CO2 그림자가격')
    a2.set_ylim(0, 400); a2.set_ylabel('CO2 그림자가격 (천원/t)')
    h1, l1 = ax[1].get_legend_handles_labels(); h2, l2 = a2.get_legend_handles_labels()
    ax[1].legend(h1 + h2, l1 + l2, loc='lower left', fontsize=8)
    ax[1].set_title('(b) 총비용과 CO2 그림자가격')
    save(fig, name)


mix_fig(SET1, 'fig1_set1.png', '일반 수요 감축률')
mix_fig(SET2, 'fig2_set2.png', '')

# 그림 3: 지역별 변화 (기준 대비 GW)
show = [k for k in opt.index if (opt.loc[k] - opt.loc[k, 'base']).abs().max() > 0.5]
dd = opt.loc[show].sub(opt.loc[show, 'base'], axis=0).drop(columns='base')
dd = dd.loc[sorted(show, key=lambda k: dd.loc[k].min())]
fig, ax = plt.subplots(figsize=(9.5, 3.4))
im = ax.imshow(dd.values, cmap='RdBu', vmin=-35, vmax=35, aspect='auto')
ax.set_yticks(range(len(dd)), [f"{REGION_KR[k.split(':')[1]]} {'PV' if k.startswith('PV') else '풍력'}"
                               f" (기준 {opt.at[k, 'base']:.1f})" for k in dd.index], fontsize=8)
ax.set_xticks(range(dd.shape[1]), [CASE_KR[c].replace('−', '-') for c in dd.columns], rotation=30, ha='right', fontsize=8)
for i in range(dd.shape[0]):
    for j in range(dd.shape[1]):
        v = dd.values[i, j]
        ax.text(j, i, f'{v:+.1f}', ha='center', va='center', fontsize=7, color='white' if abs(v) > 20 else 'black')
fig.colorbar(im, ax=ax, shrink=0.9, label='기준 대비 변화 (GW)')
ax.set_title('지역별 재생에너지 설비용량 변화 (기준 대비, 변화 0.5 GW 이상 항목)')
save(fig, 'fig3_regional_change.png')

# 그림 4: 회피비용·재생 발전 감소
d4 = s.drop('base')
fig, ax = plt.subplots(1, 2, figsize=(10, 3.6))
cols = [C_EL if c.startswith('EL') else (C_DC if c.startswith('DC_') else (C_FAB if c.startswith('Fab') else '#555555'))
        for c in d4.index]
ax[0].scatter(d4.d_demand, d4.d_cost, c=cols, s=40, zorder=3)
for c, r in d4.iterrows():
    ax[0].annotate(r.plabel, (r.d_demand, r.d_cost), textcoords='offset points', xytext=OFF.get(c, (4, -9)), fontsize=7)
xx = np.linspace(0, 135, 10)
ax[0].plot(xx, xx * 0.155, '--', color='gray', lw=1, label='기울기 155 천원/MWh')
ax[0].set_xlabel('수요 감소 (TWh)'); ax[0].set_ylabel('총비용 감소 (조원/년)'); ax[0].grid(alpha=0.3)
ax[0].legend(fontsize=7.5, loc='upper left'); ax[0].set_title('(a) 수요 감소량과 비용 감소')
ax[1].scatter(d4.d_demand, d4.d_REgen, c=cols, s=40, zorder=3)
ax[1].plot(xx, xx, '--', color='gray', lw=1, label='1:1 선')
for c, r in d4.iterrows():
    ax[1].annotate(r.plabel, (r.d_demand, r.d_REgen), textcoords='offset points', xytext=OFF.get(c, (4, -9)), fontsize=7)
ax[1].set_xlabel('수요 감소 (TWh)'); ax[1].set_ylabel('재생 발전량 감소 (TWh)'); ax[1].grid(alpha=0.3)
ax[1].legend(fontsize=7.5, loc='upper left'); ax[1].set_title('(b) 수요 감소량과 재생 발전량 감소')
save(fig, 'fig4_avoided_cost.png')

# 그림 5: 비용 부담 지표
fig, ax = plt.subplots(1, 2, figsize=(10, 3.4))
ax[0].scatter(s.d_demand, s.CO2_shadow_price, c=['k'] + cols, s=40)
for c, r in s.iterrows():
    ax[0].annotate(r.plabel, (r.d_demand, r.CO2_shadow_price), textcoords='offset points',
                   xytext={'Fab_m25': (-48, 3), 'DCFab_m25': (4, 5), 'Fab_m50': (4, -12), 'EL_m15': (4, 4),
                           'EL_m10': (-45, 3)}.get(c, (4, 3)), fontsize=7)
ax[0].set_xlabel('수요 감소 (TWh)'); ax[0].set_ylabel('CO2 그림자가격 (천원/tCO2)'); ax[0].grid(alpha=0.3)
ax[0].set_title('(a) 수요 감소와 CO2 감축 한계비용')
ax[1].scatter(s.d_demand, s.EL_price_mean, c=['k'] + cols, s=40)
for c, r in s.iterrows():
    ax[1].annotate(r.plabel, (r.d_demand, r.EL_price_mean), textcoords='offset points',
                   xytext={'DC_m50': (4, -10)}.get(c, (4, 3)), fontsize=7)
ax[1].set_xlabel('수요 감소 (TWh)'); ax[1].set_ylabel('평균 전력 한계가격 (천원/MWh)'); ax[1].grid(alpha=0.3)
ax[1].set_title('(b) 수요 감소와 평균 전력가격')
save(fig, 'fig5_burden.png')

# ───────────────────────────── 표 ─────────────────────────────
t1 = summary_table(SET1)
t2 = summary_table(SET2)
excel_tables['1_일반수요'] = t1
excel_tables['2_DC_Fab수요'] = t2
td = dem_reg.copy()
td['합계'] = td.sum(axis=1)
td.loc['합계'] = td.sum()
td.index = [REGION_KR.get(i, i) for i in td.index]
td = td.reset_index().rename(columns={'index': '지역', 'EL': '일반 (TWh)', 'DC': '데이터센터 (TWh)', 'Fab': '팹 (TWh)',
                                      '합계': '합계 (TWh)'})
excel_tables['0_지역별수요'] = td
tr = opt.copy()
tr.insert(0, '기존', exist)
tr = tr.round(1).reset_index()
tr['구분'] = tr.k.str.split(':').str[0].map({'PV': 'PV', 'WT': '풍력'})
tr['지역'] = tr.k.str.split(':').str[1].map(REGION_KR)
tr = tr.drop(columns='k').rename(columns=CASE_KR)
excel_tables['3_지역별설비(GW)'] = tr[['구분', '지역', '기존'] + list(CASE_KR.values())]
mx = mix.copy()
mx['구분'] = mx.component.map({'generator': '발전기', 'link_to_EL': '링크(→전력)'}) + ':' + mx.carrier
excel_tables['4_원별발전량(TWh)'] = mx.pivot_table(index='구분', columns='case', values='energy_TWh')[
    list(CASE_KR)].rename(columns=CASE_KR).round(1).reset_index()
excel_tables['4_원별설비(GW)'] = mx.pivot_table(index='구분', columns='case', values='capacity_GW')[
    list(CASE_KR)].rename(columns=CASE_KR).round(1).reset_index()

# ───────────────────────────── Word ─────────────────────────────
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


def B(items, style='List Bullet'):
    for it in items:
        p = doc.add_paragraph(style=style)
        if isinstance(it, tuple):
            r = p.add_run(it[0]); r.bold = True
            p.add_run(it[1])
        else:
            p.add_run(it)
        p.paragraph_format.space_after = Pt(1)


def T(df, fmt=None, font=8, caption=None):
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
                txt = '–' if pd.isna(v) else f.format(v)
            else:
                txt = str(v)
            cells[j].text = ''
            rr = cells[j].paragraphs[0].add_run(txt); rr.font.size = Pt(font)
            cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if j else WD_ALIGN_PARAGRAPH.LEFT
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def F(name, caption, width=16.5):
    doc.add_picture(figs[name], width=Cm(width))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    P(caption, italic=True, size=8.5, align=WD_ALIGN_PARAGRAPH.CENTER)


def H(text, lv):
    doc.add_heading(text, lv)


def v(c, col):
    return s.at[c, col]


COLS_SHORT = ['케이스', '총수요 (TWh)', 'PV (GW)', '풍력 (GW)', '재생 발전량 (TWh)', '출력제한률 (%)',
              'CO2 그림자가격 (천원/t)', '평균 전력가격 (천원/MWh)', '총비용 (조원)', 'MWh당 회피비용 (천원)']

# 표지·개요
doc.add_heading('PyPSA 전력계통 모형 수요 감축 민감도 분석 결과', 0)
P('2040년 재생에너지 상한 해제 시나리오: 일반·데이터센터·팹 전력수요 감축 효과', size=12)
P(f"작성일: {datetime.now():%Y-%m-%d}   |   분석 기간: 2026-10-29 ~ 2026-11-02", size=9)

H('개요', 1)
P('이전 민감도 분석이 정해진 수요에서 공급 측 조건을 바꾼 것이라면, 본 분석은 공급 측 조건을 그대로 두고 전력수요를 줄였을 때 '
  '시스템이 어떻게 변하는지를 살펴본 것이다. 기준은 현재 interface의 2040년 재생에너지 상한 해제 시나리오이며, '
  'interface 및 통합 입력 파일은 수정하지 않고 메모리 상에서 수요만 조정하였다.')
T(pd.DataFrame([
    ['기준 시나리오', '2040년 재생에너지 상한 해제 (interface.xlsx 현재 입력)'],
    ['기준 전력수요', f"일반 {v('base', 'demand_EL_TWh'):.1f} + 데이터센터 {v('base', 'demand_DC_TWh'):.1f} + "
                f"팹 {v('base', 'demand_Fab_TWh'):.1f} = {v('base', 'demand_TWh'):.1f} TWh "
                f"(EV·열·수소 부문 수요는 변경하지 않음)"],
    ['기준 결과', f"PV {v('base', 'PV_GW'):.1f} GW, 풍력 {v('base', 'WT_GW'):.1f} GW, 출력제한 "
              f"{v('base', 'curtail_%'):.1f}%, ESS {v('base', 'ESS_GWh'):.0f} GWh, 총비용 {v('base', 'cost'):.2f}조원/년, slack 없음"],
    ['CO2 제약', f"배출 상한 {co2_cap:.1f} MtCO2 (모든 케이스 동일, 구속 조건)"],
    ['수요 조정 방식', '지역별 부하의 시간대별 값 전체에 같은 비율을 곱함 → 연간 총량만 줄고 부하 형태·지역 비중은 유지'],
    ['최적화', 'PyPSA 0.31.1 + linopy, CPLEX Barrier(Crossover 생략), 8,784시간 단일연도'],
    ['비용 단위', '총비용 = 연간화 투자비 + 운전비(기존 설비 고정비는 상수로 제외), 천원 단위 입력 → 조원 환산'],
], columns=['항목', '내용']), font=8.5)
T(pd.DataFrame([
    ['1', '일반 전력수요', '{지역}_Demand_EL −5 / −10 / −15%', '수요 1 MWh 감축당 약 15.3~15.7만 원 비용 절감, 풍력부터 감소'],
    ['2', '데이터센터·팹 수요', 'DC, Fab, 동시 각각 −25 / −50%', '같은 감축량이면 일반 수요와 효과가 거의 같음. '
                                                       '동시 −50%에서 신규 풍력 0'],
], columns=['No.', '세트', '변경 조건', '핵심 결과']), font=8, caption='표 0-1. 분석 세트 요약')
T(td, font=8, caption='표 0-2. 기준 시나리오의 지역별 전력수요 (TWh)')
F('fig0_demand_profile.png', '그림 0. 지역별 수요 구성(a), 시간대별(b)·월별(c) 평균 부하 형태 (전국 합계, 평균=1)', width=17)
H('핵심 메시지', 2)
B([('수요 감축은 재생에너지 신규 설치를 거의 1:1로 줄인다. ', 'CO2 상한이 고정되어 가스 발전(약 158.5 TWh)과 원자력(약 259 TWh)은 '
    '모든 케이스에서 같고, 줄어든 수요만큼 재생 발전량이 감소한다.'),
   ('수요 1 MWh를 줄이면 시스템 비용이 약 15만 원 줄어든다. ', '9개 감축 케이스의 회피비용은 '
    f"{s.avoided.min():.0f}~{s.avoided.max():.0f} 천원/MWh로 감축 대상·규모와 무관하게 거의 일정하다."),
   ('어떤 수요를 줄이느냐보다 얼마나 줄이느냐가 중요하다. ', '일반 −15%(85.9 TWh)와 팹 −50%(85.5 TWh)는 총비용·설비 구성이 거의 같다. '
    '다만 감축 지역에 따라 줄어드는 설비의 입지가 달라진다.'),
   ('수요 감축은 탄소중립 비용 부담을 낮춘다. ', f"CO2 그림자가격 {v('base', 'CO2_shadow_price'):.0f} → "
    f"{v('DCFab_m50', 'CO2_shadow_price'):.0f} 천원/t, 평균 전력가격 {v('base', 'EL_price_mean'):.0f} → "
    f"{v('DCFab_m50', 'EL_price_mean'):.0f} 천원/MWh (DC+Fab −50%). 같은 배출 목표를 더 낮은 한계비용으로 달성한다."),
   ('출력제한률과 ESS는 변하지 않는다. ', '수요와 함께 재생 설비도 줄어 출력제한률은 11~12%로 유지되고, '
    'ESS는 모든 케이스에서 기존 용량(107 GWh)에 머문다.')])

# 세트 1
doc.add_page_break()
H('세트 1. 일반 전력수요(Demand_EL) 감축', 1)
H('분석 목적 및 설정', 2)
B(['효율 향상·수요관리 등으로 일반 전력수요가 줄어들 때 재생에너지 보급 규모, 비용, CO2 감축 부담이 어떻게 변하는지 검토',
   f"11개 지역 {{지역}}_Demand_EL(기준 {v('base', 'demand_EL_TWh'):.1f} TWh)에 0.95 / 0.90 / 0.85 배율 적용",
   '데이터센터·팹·EV·열·수소 수요와 공급 측 조건은 기준과 동일'])
T(t1[COLS_SHORT], fmt=FMT, font=7.5, caption='표 1. 일반 전력수요 감축률별 결과')
F('fig1_set1.png', '그림 1. 일반 수요 감축률별 재생에너지 설비용량(a)과 총비용·CO2 그림자가격(b)')
H('주요 결과 및 해석', 2)
B([('재생 설비 감소: ', f"−15%(−{v('EL_m15', 'd_demand'):.1f} TWh)에서 PV {v('base', 'PV_GW'):.0f}→{v('EL_m15', 'PV_GW'):.0f} GW, "
    f"풍력 {v('base', 'WT_GW'):.0f}→{v('EL_m15', 'WT_GW'):.0f} GW로 합계 "
    f"{v('base', 'RE_GW') - v('EL_m15', 'RE_GW'):.0f} GW 감소. 감축 1단계(−5%)마다 재생 약 18~22 GW씩 줄어듦"),
   ('풍력이 먼저 줄어듦: ', f"서울경기 풍력이 {opt.at['WT:GGD', 'base']:.1f}→{opt.at['WT:GGD', 'EL_m15']:.1f} GW로 대부분 빠짐. "
    '모형 내에서 가장 비싼(이용률 대비 비용이 높은) 재생 설비부터 줄어드는 것. PV는 충남·광주전남·부산울산 중심으로 감소'),
   ('발전믹스: ', '가스(158.5 TWh)·원자력(259 TWh)·석탄(0)은 불변. 재생 발전량이 수요 감소량과 거의 같은 폭으로 감소 '
    f"({v('EL_m15', 'd_REgen'):.1f} TWh vs 수요 {v('EL_m15', 'd_demand'):.1f} TWh)"),
   ('비용: ', f"총비용 {v('base', 'cost'):.2f}→{v('EL_m15', 'cost'):.2f}조원(−{v('EL_m15', 'd_cost'):.1f}조원). "
    f"MWh당 회피비용 {v('EL_m05', 'avoided'):.0f}→{v('EL_m15', 'avoided'):.0f} 천원으로, 감축 폭이 커질수록 "
    '조금씩 낮아짐(비싼 설비부터 빠지므로 한계 절감 효과가 체감)'),
   ('CO2 부담: ', f"그림자가격 {v('base', 'CO2_shadow_price'):.0f}→{v('EL_m15', 'CO2_shadow_price'):.0f} 천원/t(−23%), "
    f"평균 전력가격 {v('base', 'EL_price_mean'):.1f}→{v('EL_m15', 'EL_price_mean'):.1f} 천원/MWh"),
   ('출력제한: ', '11.7% → 10.9~11.1%로 소폭 하락. 수요와 함께 설비도 줄어들어 출력제한 구조는 유지')])
H('시사점 (시나리오 반영)', 2)
B(['일반 수요 전망의 불확실성(±5~15%)은 재생 보급 규모를 20~65 GW 범위로 바꿀 수 있으므로, 보급목표 제시 시 수요 전망과 함께 제시 필요',
   '수요관리·효율 향상의 가치는 "MWh당 약 15만 원의 시스템 비용 회피"로 정량화 가능 → 효율 투자 비용과 직접 비교할 수 있는 기준값',
   '감축분은 주로 수도권 풍력·남부 PV의 신규 설치 회피로 나타나므로, 입지 갈등이 큰 설비의 수요를 줄이는 효과도 있음'])

# 세트 2
doc.add_page_break()
H('세트 2. 데이터센터·팹 수요 감축', 1)
H('분석 목적 및 설정', 2)
B(['데이터센터(DC)·반도체 팹(Fab) 등 대규모 신규 수요의 전망 불확실성이 전력시스템에 미치는 영향 검토',
   f"DC(기준 {v('base', 'demand_DC_TWh'):.1f} TWh: 강원·광주전남·충남·부산울산·대구경북·전북), "
   f"Fab(기준 {v('base', 'demand_Fab_TWh'):.1f} TWh: 서울경기·충북·광주전남·충남·전북)",
   'DC만 / Fab만 / 동시에 각각 −25%, −50%. 재생에너지 전용 DC·Fab 수요(_RE)는 기준에서 모두 0이므로 대상 아님',
   f"부하 형태: 일반 수요 부하율 {lf['EL']:.2f}, DC {lf['DC']:.2f}, Fab {lf['Fab']:.2f}(그림 0)"])
T(t2[COLS_SHORT], fmt=FMT, font=7.5, caption='표 2. 데이터센터·팹 수요 감축 케이스별 결과')
F('fig2_set2.png', '그림 2. 케이스별 재생에너지 설비용량(a)과 총비용·CO2 그림자가격(b)')
H('주요 결과 및 해석', 2)
B([('감축량이 결과를 결정: ', 'DC −50%(42.6 TWh)와 Fab −25%(42.7 TWh)는 감축량이 같고 결과도 거의 같음 '
    f"(총비용 {v('DC_m50', 'cost'):.2f} vs {v('Fab_m25', 'cost'):.2f}조원, 재생 {v('DC_m50', 'RE_GW'):.0f} vs "
    f"{v('Fab_m25', 'RE_GW'):.0f} GW). 일반 −15%와 Fab −50%도 같은 관계"),
   ('동시 −50%: ', f"수요 −{v('DCFab_m50', 'd_demand'):.0f} TWh로 재생 {v('base', 'RE_GW'):.0f}→{v('DCFab_m50', 'RE_GW'):.0f} GW, "
    f"신규 풍력이 0이 됨(풍력 {v('DCFab_m50', 'WT_GW'):.0f} GW = 기존 설비). 총비용 −{v('DCFab_m50', 'd_cost'):.1f}조원"),
   ('지역 효과: ', 'Fab 수요의 57%가 서울경기에 있으나 서울경기 PV는 이미 기존 설비 수준이므로, Fab 감축 시에는 서울경기로 '
    f"전력을 보내던 충남 PV({opt.at['PV:CND', 'base']:.1f}→{opt.at['PV:CND', 'Fab_m50']:.1f} GW)와 충북·서울경기 풍력이 더 줄어듦. "
    'DC 감축은 충남·광주전남·부산울산 PV를 주로 줄임. 서울경기 풍력은 모든 케이스에서 가장 먼저 감소'),
   ('부하 형태의 영향은 작음: ', 'DC·Fab은 일반 수요와 부하 형태가 다르지만(그림 0), MWh당 회피비용은 '
    f"{s.loc[SET2[1:], 'avoided'].min():.0f}~{s.loc[SET2[1:], 'avoided'].max():.0f} 천원으로 일반 수요 감축"
    f"({s.loc[SET1[1:], 'avoided'].min():.0f}~{s.loc[SET1[1:], 'avoided'].max():.0f} 천원)과 차이가 작음. "
    '원자력·가스 발전량이 고정된 상태에서 재생 설비와 기존 ESS·송전망으로 조정되기 때문으로 해석'),
   ('전력가격에서는 차이: ', f"감축량이 비슷한 일반 −15%와 Fab −50%를 비교하면 CO2 그림자가격은 거의 같지만"
    f"({v('EL_m15', 'CO2_shadow_price'):.0f} vs {v('Fab_m50', 'CO2_shadow_price'):.0f}), 평균 전력가격은 일반 수요 감축이 더 낮음"
    f"({v('EL_m15', 'EL_price_mean'):.1f} vs {v('Fab_m50', 'EL_price_mean'):.1f} 천원/MWh). 첨두가 뚜렷한 일반 수요를 줄이면 "
    '가격이 높은 시간대의 수요가 더 많이 줄기 때문(그림 4b)'),
   ('출력제한·ESS: ', '출력제한률 11.1~11.8%, ESS 107 GWh(기존)로 거의 불변')])
H('시사점 (시나리오 반영)', 2)
B(['데이터센터·팹 수요 전망은 규모가 커서(합계 256 TWh, 전체의 31%) 재생 보급 규모를 좌우하는 핵심 변수 → '
   '고·저 수요 시나리오를 별도로 구성할 것을 권장',
   'DC+Fab 수요가 절반이면 연간 약 19.5조원의 시스템 비용이 줄고, 신규 풍력 수요가 사라짐 → 풍력 보급목표는 대규모 수요 전망과 연계해 설정 필요',
   '대규모 수요의 입지(지역)는 재생 설비 입지에 영향을 주므로, 수요 입지 분산 정책과 재생 입지 정책을 함께 검토할 수 있음'])

# 세트 간 비교·비용 부담
doc.add_page_break()
H('종합: 수요 감축의 비용 효과와 탄소중립 비용 부담', 1)
F('fig4_avoided_cost.png', '그림 3. 수요 감소량과 총비용 감소(a), 재생 발전량 감소(b) — 파랑: 일반, 주황: DC, 보라: 팹, 회색: DC+Fab')
t5 = s.drop('base')[['label', 'd_demand', 'd_cost', 'avoided', 'd_REgen', 'CO2_shadow_price', 'EL_price_mean']].rename(columns={
    'label': '케이스', 'd_demand': '수요 감소 (TWh)', 'd_cost': '비용 감소 (조원)', 'avoided': 'MWh당 회피비용 (천원)',
    'd_REgen': '재생 발전 감소 (TWh)', 'CO2_shadow_price': 'CO2 그림자가격 (천원/t)', 'EL_price_mean': '평균 전력가격 (천원/MWh)'})
excel_tables['5_회피비용'] = t5
T(t5, fmt=FMT, font=8, caption='표 3. 케이스별 수요 감축 회피비용 (기준 대비)')
F('fig5_burden.png', '그림 4. 수요 감소량과 CO2 그림자가격(a)·평균 전력 한계가격(b)')
H('비용 부담 감소의 의미', 2)
B([('CO2 그림자가격 = 탄소 감축의 한계비용: ', '배출 상한을 1톤 더 강화할 때 늘어나는 시스템 비용이다. 수요가 줄면 같은 배출 상한을 '
    '맞추기 위해 비싼 재생 설비를 덜 지어도 되므로 한계비용이 내려간다 '
    f"({v('base', 'CO2_shadow_price'):.0f} → {v('DCFab_m50', 'CO2_shadow_price'):.0f} 천원/t, −"
    f"{(1 - v('DCFab_m50', 'CO2_shadow_price') / v('base', 'CO2_shadow_price')) * 100:.0f}%)."),
   ('정책적 해석: ', '수요 감축은 그 자체로 배출을 줄이는 것이 아니라(배출량은 상한에 고정), 같은 감축 목표를 더 싸게 달성하게 해 준다. '
    '따라서 수요관리 정책은 "탄소중립 이행 비용을 낮추는 수단"으로 제시할 수 있다.'),
   ('추가 감축 여력: ', '그림자가격이 낮아진다는 것은 같은 비용으로 더 강한 배출 목표를 설정할 여지가 생긴다는 의미이기도 하다. '
    '예: 수요 감축 케이스에서 배출 상한을 강화했을 때의 비용 비교는 후속 분석 과제'),
   ('전력가격: ', f"평균 한계가격은 {v('base', 'EL_price_mean'):.0f} → {v('DCFab_m50', 'EL_price_mean'):.0f} 천원/MWh로 "
    '약 10% 하락. 기존 수요자 입장에서도 요금 부담 완화로 이어질 수 있음(단, 모형의 한계가격은 실제 요금과 다름)')])
H('회피비용 활용 방안', 2)
B(['MWh당 약 15만 원(153~159 천원)은 "1 MWh의 수요를 줄였을 때 회피되는 신규 재생설비·운영비용"으로, '
   '효율 향상 투자·수요관리 프로그램의 비용효과 판단 기준으로 활용 가능',
   '감축 규모가 커질수록 회피비용이 소폭 낮아지는데(비싼 설비부터 빠지므로), 대규모 감축 시에는 이 체감 효과를 반영해야 함',
   '회피비용은 CO2 상한이 구속되는 조건에서의 값이다. 상한이 느슨해지면 가스 발전이 한계 공급원이 되어 회피비용 구조가 달라질 수 있음'])
H('지역별 설비 변화', 2)
F('fig3_regional_change.png', '그림 5. 지역별 재생에너지 설비용량 변화 (기준 대비 GW, 변화가 없는 지역·기술은 생략)', width=17)
B(['줄어드는 설비는 서울경기 풍력 → 충남·광주전남·부산울산 PV 순. 나머지 지역의 PV·풍력은 기존 설비(하한)에 머물러 변화 없음',
   '서울경기 풍력은 기준에서 34 GW로 기술적 잠재량(44 GW)의 77%까지 설치되는데, 수요가 줄면 가장 먼저 사라짐 → '
   '수도권 풍력 대규모 보급은 높은 수요 전망을 전제로 한 결과임을 보고서에 명시 필요',
   '반대로 충북 풍력은 일반 수요 감축 시 소폭 증가(14.8→16.0 GW)하는데, 서울경기 풍력 감소분 일부를 대체하는 재배치 효과'])

# 부록
doc.add_page_break()
H('부록. 방법론 및 결과 파일', 1)
H('A. 구현', 2)
B([('수요 조정: ', '네트워크 생성 후 대상 부하의 p_set(정적값·시계열 모두)에 배율을 곱함. '
    '부하 이름은 {지역}_Demand_EL / {지역}_Demand_DC / {지역}_Demand_Fab (입력의 DC 부하 이름 끝 공백 포함 처리)'),
   ('회피비용: ', 'MWh당 회피비용 = (기준 총비용 − 케이스 총비용) ÷ (기준 수요 − 케이스 수요)'),
   ('솔버: ', 'CPLEX Barrier convergetol 1e-6. Fab −50%, DC+Fab −25/−50%는 1e-6에서 수렴 판정 실패로 1e-5로 재실행'
    '(이전 송전망 분석에서 같은 케이스를 1e-5와 1e-6으로 풀었을 때 총비용 차이 0.001% 이내 확인)'),
   ('결과 지표: ', '평균 전력가격 = 11개 지역 전력 버스 한계가격의 단순 평균(수요 가중 아님)')])
H('B. 결과 파일', 2)
T(pd.DataFrame([
    ['기준, 세트 1, 세트 2(DC −25/−50, Fab −25)', DIRS[0], 'run_demand_sweep.py'],
    ['세트 2(Fab −50, DC+Fab −25/−50)', DIRS[1], 'run_demand_sweep.py --only ... (convergetol 1e-5)'],
    ['보고서·그림·표', OUT, 'make_demand_report.py'],
], columns=['내용', '결과 폴더', '실행 스크립트']), font=8)
H('C. 유의사항', 2)
B(['모든 결과는 2040년 단일연도·현재 입력조건 기준이며, 공급 측(원자력·석탄·가스 설비, CO2 상한)은 고정',
   '총비용은 기존 설비 고정비를 제외한 값으로, 케이스 간 "차이"로 해석해야 함',
   '수요 감축에 드는 비용(효율 투자, 산업 유치 포기 등)은 모형에 포함되지 않음 → 회피비용과 별도로 비교 필요',
   '수요는 연간 총량만 비례 축소하였으며, 수요반응(시간대 이동)은 반영하지 않음'])

docx_path = os.path.join(OUT, '수요감축_민감도분석_보고서.docx')
doc.save(docx_path)
xlsx_path = os.path.join(OUT, '수요감축_민감도분석_표.xlsx')
with pd.ExcelWriter(xlsx_path) as wr:
    for name, df in excel_tables.items():
        df.to_excel(wr, sheet_name=name[:31], index=False)
print('완료:', docx_path, xlsx_path, sep='\n')
print(s[['label', 'avoided']].to_string())
print({k: round(x, 3) for k, x in lf.items()})
print(dem_reg.round(1).to_string())
