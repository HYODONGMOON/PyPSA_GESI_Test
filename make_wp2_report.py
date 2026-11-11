"""
반도체 수요 지역 재배치(WP2) 비교 및 추가 민감도 분석 보고서(Word) + 그림 + 표 생성

선행: python compare_wp2.py  (reports/wp2_compare_<날짜>/national.csv, regional.csv, corridor.csv)
사용법:
  python make_wp2_report.py
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

D = sorted(glob.glob('reports/wp2_compare_*'))[-1]
OUT = D
nat = pd.read_csv(f'{D}/national.csv', index_col=0)
reg = pd.read_csv(f'{D}/regional.csv')
corr = pd.read_csv(f'{D}/corridor.csv')
AVOID = 155.0  # 천원/MWh, 수요 감축 민감도 분석의 MWh당 회피비용

REGION_KR = {'BSN': '부산울산', 'CBD': '충북', 'CND': '대전세종충남', 'GBD': '대구경북', 'GGD': '서울경기',
             'GND': '경남', 'GWD': '강원', 'ICN': '인천', 'JBD': '전북', 'JJD': '제주', 'JND': '광주전남'}
REG = ['GGD', 'ICN', 'GWD', 'CBD', 'CND', 'JBD', 'GBD', 'BSN', 'GND', 'JND', 'JJD']
WP2 = {'2035': 'WP2fix_2035', '2040': 'WP2_2040'}
YEARS = ('2035', '2040')
C_BASE, C_WP2, C_PV, C_WT, C_SL = '#9E9E9E', '#8E5EA2', '#F2B134', '#3C8DBC', '#C0392B'


def kr_corr(c):
    a, b = c.split('-')
    return f'{REGION_KR[a]}–{REGION_KR[b]}'


def N(s, k):
    return float(nat.at[k, s])


def R(s):
    return reg[reg.scenario == s].set_index('region').reindex(REG)


def extra(s):
    r = R(s)
    fab = r.Fab_TWh
    return {'metro_import': -(r.loc[['GGD', 'ICN'], 'net_export_TWh'].sum()),
            'fab_re_ss': (fab * r['RE_selfsuff_%']).sum() / fab.sum(),
            'cbd_wt_pot': r.at['CBD', 'WT_GW'] / TECH_POTENTIAL_GW['WT']['CBD'] * 100,
            'hvdc': corr[(corr.scenario == s) & (corr.type == 'HVDC')].flow_abs_TWh.sum()}


EX = {s: extra(s) for s in nat.columns}


def pair_rows(a, b, y):
    dfab = N(b, 'demand_Fab_TWh') - N(a, 'demand_Fab_TWh')
    dcost = N(b, 'cost_excl_penalty_조') - N(a, 'cost_excl_penalty_조')
    return dfab, dcost, dcost - dfab * AVOID / 1e3


def ind_table(cols):
    rows = [
        ('시스템 비용 (조원/년)', 'cost_excl_penalty_조', '{:.2f}'),
        ('MWh당 시스템 비용 (천원)', 'cost_per_MWh_천원', '{:.1f}'),
        ('PV (GW)', 'PV_GW', '{:.1f}'), ('풍력 (GW)', 'WT_GW', '{:.1f}'),
        ('출력제한률 (%)', 'curtail_%', '{:.1f}'), ('공급부족 slack (TWh)', 'slack_TWh', '{:.1f}'),
        ('CO2 그림자가격 (천원/t)', 'CO2_shadow_천원/t', '{:,.0f}'),
        ('반도체 가중 전력가격 (천원/MWh)', 'price_Fab_loadw', '{:,.1f}'),
        ('송전 수송거리 (TWh·km)', 'grid_TWhkm', '{:,.0f}'), ('송전 혼잡시간 합계 (h)', 'grid_cong_h_sum', '{:,.0f}'),
        ('혼잡비용 (조원)', 'grid_rent_조', '{:.2f}'), ('송전 증설 (GW)', 'grid_added_GW', '{:.1f}')]
    out = []
    for lab, k, f in rows:
        out.append([lab] + [f.format(N(s, k)) for s, _ in cols])
    for lab, k, f in (('수도권 순수입 (TWh)', 'metro_import', '{:.1f}'), ('반도체 입지 재생 자급률 (%)', 'fab_re_ss', '{:.1f}'),
                      ('충북 풍력 / 잠재량 (%)', 'cbd_wt_pot', '{:.0f}'), ('HVDC 수송량 (TWh)', 'hvdc', '{:.1f}')):
        out.append([lab] + [f.format(EX[s][k]) for s, _ in cols])
    return pd.DataFrame(out, columns=['지표'] + [c for _, c in cols])


figs = {}


def save(fig, name):
    p = os.path.join(OUT, name)
    fig.tight_layout()
    fig.savefig(p, dpi=200)
    plt.close(fig)
    figs[name] = p


# ───────────── 그림 ─────────────
# fig1 반도체 배분
fig, ax = plt.subplots(1, 2, figsize=(10, 3.3), sharey=True)
for k, y in enumerate(YEARS):
    A, B = R(f'OPEN_{y}').Fab_TWh, R(WP2[y]).Fab_TWh
    x = np.arange(len(REG))
    ax[k].bar(x - 0.2, A, 0.4, color=C_BASE, label='기준(정부 발표)')
    ax[k].bar(x + 0.2, B, 0.4, color=C_WP2, label='재배치(기업 설문)')
    ax[k].set_xticks(x, [REGION_KR[r] for r in REG], rotation=45, ha='right', fontsize=8)
    ax[k].set_title(f'({chr(97 + k)}) {y}년 반도체 수요 (TWh)'); ax[k].legend(fontsize=8); ax[k].grid(axis='y', alpha=0.3)
save(fig, 'fig1_fab_allocation.png')


def regional_change_fig(pairs, name, titles):
    fig, ax = plt.subplots(1, 3, figsize=(12, 3.6), sharey=True)
    yy = np.arange(len(REG))[::-1]
    for k, (a, b) in enumerate(pairs):
        A, B = R(a), R(b)
        ax[k].barh(yy + 0.2, B.PV_GW - A.PV_GW, 0.4, color=C_PV, label='PV')
        ax[k].barh(yy - 0.2, B.WT_GW - A.WT_GW, 0.4, color=C_WT, label='풍력')
        ax[k].axvline(0, color='k', lw=0.6); ax[k].set_yticks(yy, [REGION_KR[r] for r in REG], fontsize=8)
        ax[k].set_xlabel('재배치 - 기준 (GW)'); ax[k].set_title(titles[k]); ax[k].legend(fontsize=7.5, loc='lower right')
        ax[k].grid(axis='x', alpha=0.3)
    for j, (a, b) in enumerate(pairs):
        d = -(R(b).net_export_TWh - R(a).net_export_TWh)
        ax[2].barh(yy + (0.2 if j == 0 else -0.2), d, 0.4, color='#7FA7C9' if j == 0 else '#1F4E79', label=YEARS[j] + '년')
    ax[2].axvline(0, color='k', lw=0.6); ax[2].set_xlabel('순수입 변화 (TWh, +는 수입 증가)')
    ax[2].set_title('(c) 지역 순수입량 변화'); ax[2].legend(fontsize=7.5); ax[2].grid(axis='x', alpha=0.3)
    save(fig, name)


regional_change_fig([('OPEN_2035', WP2['2035']), ('OPEN_2040', WP2['2040'])], 'fig2_regional_change.png',
                    ['(a) 2035년 재생설비 변화', '(b) 2040년 재생설비 변화'])
regional_change_fig([('OPENcap_2035', 'WP2cap_2035'), ('OPENcap_2040', 'WP2cap_2040')], 'fig6_cap_regional.png',
                    ['(a) 2035년 재생설비 변화 (잠재량 20% 상한)', '(b) 2040년 재생설비 변화 (잠재량 20% 상한)'])


def grid_fig(a, b, name, title):
    A = corr[corr.scenario == a].set_index(['type', 'corridor'])
    B = corr[corr.scenario == b].set_index(['type', 'corridor'])
    d = pd.DataFrame({'bf': A.flow_abs_TWh, 'nf': B.flow_abs_TWh, 'bh': A.cong_h, 'nh': B.cong_h})
    d = d[d.bf > 0.5]
    d = d.loc[(d.nf - d.bf).abs().sort_values().index]
    lab = [f"{'[HVDC] ' if t == 'HVDC' else ''}{kr_corr(c)}" for t, c in d.index]
    fig, ax = plt.subplots(1, 2, figsize=(11, 4.6), sharey=True)
    yy = np.arange(len(d))
    ax[0].barh(yy + 0.2, d.bf, 0.4, color=C_BASE, label='기준')
    ax[0].barh(yy - 0.2, d.nf, 0.4, color='#1F4E79', label='재배치')
    ax[0].set_yticks(yy, lab, fontsize=8); ax[0].set_xlabel('연간 수송량 (TWh, 양방향 합)')
    ax[0].set_title(f'(a) {title} 구간별 송전 수송량'); ax[0].legend(fontsize=8); ax[0].grid(axis='x', alpha=0.3)
    ax[1].barh(yy + 0.2, d.bh, 0.4, color=C_BASE, label='기준')
    ax[1].barh(yy - 0.2, d.nh, 0.4, color=C_SL, label='재배치')
    ax[1].set_xlabel('혼잡시간 (h/년)'); ax[1].set_title(f'(b) {title} 구간별 혼잡시간'); ax[1].legend(fontsize=8)
    ax[1].grid(axis='x', alpha=0.3)
    save(fig, name)


grid_fig('OPEN_2040', 'WP2_2040', 'fig3_grid_2040.png', '2040년')

# fig4 REF vs OPEN
r1, r2 = R('REF_2040'), R('OPEN_2040')
fig, ax = plt.subplots(1, 2, figsize=(10, 3.4))
x = np.arange(len(REG))
ax[0].bar(x - 0.2, r1.PV_GW + r1.WT_GW, 0.4, color='#BDBDBD', label='REF(상한)')
ax[0].bar(x + 0.2, r2.PV_GW + r2.WT_GW, 0.4, color='#2E7D32', label='상한 해제')
ax[0].set_xticks(x, [REGION_KR[r] for r in REG], rotation=45, ha='right', fontsize=8)
ax[0].set_ylabel('PV+풍력 (GW)'); ax[0].set_title('(a) 2040년 지역별 재생설비'); ax[0].legend(fontsize=8)
ax[1].bar(x, r1.slack_TWh, color=C_SL)
for i, v in enumerate(r1.slack_TWh):
    if v > 0.5:
        ax[1].text(i, v + 1, f'{v:.0f}', ha='center', fontsize=8)
ax[1].set_xticks(x, [REGION_KR[r] for r in REG], rotation=45, ha='right', fontsize=8)
ax[1].set_ylabel('TWh'); ax[1].set_title('(b) 2040년 REF의 지역별 공급부족(slack)')
save(fig, 'fig4_ref_vs_open_2040.png')

# fig5 REF × 재배치: 지역별 slack
fig, ax = plt.subplots(1, 2, figsize=(10, 3.4), sharey=True)
for k, y in enumerate(YEARS):
    A, B = R(f'REF_{y}').slack_TWh, R(f'REFWP2_{y}').slack_TWh
    ax[k].bar(x - 0.2, A, 0.4, color=C_BASE, label=f'기준배치 (계 {A.sum():.0f})')
    ax[k].bar(x + 0.2, B, 0.4, color=C_SL, label=f'재배치 (계 {B.sum():.0f})')
    for i in range(len(REG)):
        if B.iloc[i] > 0.5:
            ax[k].text(i + 0.2, B.iloc[i] + 0.8, f'{B.iloc[i]:.0f}', ha='center', fontsize=7.5)
    ax[k].set_xticks(x, [REGION_KR[r] for r in REG], rotation=45, ha='right', fontsize=8)
    ax[k].set_title(f'({chr(97 + k)}) {y}년 REF 조건 지역별 공급부족 (TWh)'); ax[k].legend(fontsize=8)
    ax[k].grid(axis='y', alpha=0.3)
save(fig, 'fig5_ref_wp2_slack.png')

# fig7 송전 증설 구간
fig, ax = plt.subplots(1, 2, figsize=(11, 3.8), sharex=True)
for k, y in enumerate(YEARS):
    A = corr[corr.scenario == f'OPENcapgrid_{y}'].set_index(['type', 'corridor']).added_GW
    B = corr[corr.scenario == f'WP2capgrid_{y}'].set_index(['type', 'corridor']).added_GW
    d = pd.DataFrame({'a': A, 'b': B})
    d = d[d.max(axis=1) > 0.1].sort_values('a')
    yy = np.arange(len(d))
    ax[k].barh(yy + 0.2, d.a, 0.4, color=C_BASE, label=f'기준배치 (계 {A.sum():.1f} GW)')
    ax[k].barh(yy - 0.2, d.b, 0.4, color='#1F4E79', label=f'재배치 (계 {B.sum():.1f} GW)')
    for i, (va, vb) in enumerate(zip(d.a, d.b)):
        ax[k].text(max(va, vb) + 0.3, i, f'{va:.1f} → {vb:.1f}', va='center', fontsize=7)
    ax[k].set_yticks(yy, [kr_corr(c) for _, c in d.index], fontsize=8)
    ax[k].set_xlabel('증설 용량 (GW)'); ax[k].set_title(f'({chr(97 + k)}) {y}년 필요 구간 송전 증설')
    ax[k].legend(fontsize=7.5, loc='lower right'); ax[k].grid(axis='x', alpha=0.3); ax[k].set_xlim(0, 25)
save(fig, 'fig7_grid_expansion.png')

# fig8 종합: 조건별 재배치 순효과, 수도권 순수입, 혼잡시간
conds = [('상한 해제', 'OPEN_{y}', '{wp2}'), ('잠재량 20%', 'OPENcap_{y}', 'WP2cap_{y}'),
         ('잠재량 20%\n+송전 확장', 'OPENcapgrid_{y}', 'WP2capgrid_{y}')]
summ = []
for y in YEARS:
    for lab, a, b in conds:
        a_, b_ = a.format(y=y), b.format(y=y, wp2=WP2[y])
        dfab, dcost, net = pair_rows(a_, b_, y)
        summ.append({'연도': y, '조건': lab.replace('\n', ' '), '비용 변화 (조원)': dcost, '반도체 수요 차이 (TWh)': dfab,
                     '재배치 순효과 (조원)': net,
                     '수도권 순수입 기준 (TWh)': EX[a_]['metro_import'], '수도권 순수입 재배치 (TWh)': EX[b_]['metro_import'],
                     '혼잡시간 기준 (h)': N(a_, 'grid_cong_h_sum'), '혼잡시간 재배치 (h)': N(b_, 'grid_cong_h_sum'),
                     '반도체 가중가격 기준': N(a_, 'price_Fab_loadw'), '반도체 가중가격 재배치': N(b_, 'price_Fab_loadw')})
summ = pd.DataFrame(summ)
fig, ax = plt.subplots(1, 3, figsize=(12, 3.5))
xl = np.arange(3)
for j, y in enumerate(YEARS):
    s = summ[summ.연도 == y]
    ax[0].bar(xl + (j - 0.5) * 0.35, s['재배치 순효과 (조원)'], 0.35, color=['#7FA7C9', '#1F4E79'][j], label=f'{y}년')
    ax[1].bar(xl + (j - 0.5) * 0.35, s['수도권 순수입 재배치 (TWh)'] - s['수도권 순수입 기준 (TWh)'], 0.35,
              color=['#7FA7C9', '#1F4E79'][j], label=f'{y}년')
    ax[2].bar(xl + (j - 0.5) * 0.35, 100 * (s['혼잡시간 재배치 (h)'] / s['혼잡시간 기준 (h)'] - 1), 0.35,
              color=['#7FA7C9', '#1F4E79'][j], label=f'{y}년')
for a_, t, yl in ((ax[0], '(a) 재배치 순효과 (수요 차이 보정)', '조원/년'),
                  (ax[1], '(b) 수도권 순수입 변화', 'TWh'), (ax[2], '(c) 송전 혼잡시간 변화', '%')):
    a_.set_xticks(xl, [c[0] for c in conds], fontsize=8); a_.axhline(0, color='k', lw=0.6)
    a_.set_title(t); a_.set_ylabel(yl); a_.legend(fontsize=8); a_.grid(axis='y', alpha=0.3)
save(fig, 'fig8_summary.png')

# ───────────── Excel ─────────────
LAB = {'REF_2035': '2035 REF·기준배치', 'REFWP2_2035': '2035 REF·재배치', 'OPEN_2035': '2035 상한해제·기준배치',
       'WP2_2035': '2035 상한해제·재배치(원본)', 'WP2fix_2035': '2035 상한해제·재배치', 'OPENcap_2035': '2035 잠재량20%·기준배치',
       'WP2cap_2035': '2035 잠재량20%·재배치', 'OPENcapgrid_2035': '2035 잠재량20%+송전·기준배치',
       'WP2capgrid_2035': '2035 잠재량20%+송전·재배치'}
LAB.update({k.replace('2035', '2040'): v.replace('2035', '2040') for k, v in LAB.items()})
with pd.ExcelWriter(os.path.join(OUT, 'WP2_비교_민감도_표.xlsx')) as wr:
    nat.rename(columns=LAB).to_excel(wr, sheet_name='전국지표')
    summ.to_excel(wr, sheet_name='재배치효과_요약', index=False)
    for col in ('Fab_TWh', 'load_TWh', 'PV_GW', 'WT_GW', 'slack_TWh', 'net_export_TWh', 'RE_selfsuff_%', 'price_loadw'):
        reg.pivot_table(index='region', columns='scenario', values=col)[list(nat.columns)].rename(
            columns=LAB, index=REGION_KR).to_excel(wr, sheet_name=f'지역_{col}'[:31])
    for col in ('flow_abs_TWh', 'cong_h', 'added_GW', 'rent_조'):
        corr.pivot_table(index=['type', 'corridor'], columns='scenario', values=col)[list(nat.columns)].rename(
            columns=LAB).to_excel(wr, sheet_name=f'송전_{col}'[:31])

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


def T(df, font=8, caption=None):
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
            txt = f'{v:,.2f}' if isinstance(v, (float, np.floating)) else str(v)
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


def r_(s, r, col):
    return float(R(s).at[r, col])


s35, s40 = summ[summ.연도 == '2035'].set_index('조건'), summ[summ.연도 == '2040'].set_index('조건')

# 표지·개요
doc.add_heading('반도체 수요 지역 재배치에 따른 전력시스템 영향 분석', 0)
P('정부 발표 기준 배치 vs 주요 기업 설문 기반 재배치 — 2035·2040년, 재생 상한 설정/해제 및 추가 민감도', size=12)
P(f'작성일: {datetime.now():%Y-%m-%d}   |   모형: PyPSA 0.31.1 + CPLEX, 8,784시간 단일연도', size=9)

H('개요', 1)
P('반도체(팹) 전력수요의 총량은 같고 지역 배분만 다른 두 배치(정부 발표 기준, 주요 기업 설문 기반 재배치)가 전력 수급, '
  '전력망, 비용에 미치는 차이를 비교하였다. 기준 결과는 재생에너지 보급 상한 설정(REF)과 해제(OPEN) 두 조건이며, '
  '재배치 결과(WP2)는 상한 해제 조건으로 계산되었다. 여기에 ① REF 상한 조건 재배치, ② 지역 잠재량 상한, '
  '③ 잠재량 상한 + 필요 구간 송전 확장의 세 가지 민감도 분석을 추가하였다.')
T(pd.DataFrame([
    ['기준 비교', 'REF(상한) / 상한 해제 × 기준배치', '상한 해제 × 재배치', '2035, 2040', '기존 결과'],
    ['민감도 1', 'REF 상한 × 기준배치', 'REF 상한 × 재배치', '2035, 2040', '재배치 네트워크에 REF 상한 적용'],
    ['민감도 2', '잠재량 20% 상한 × 기준배치', '잠재량 20% 상한 × 재배치', '2035, 2040', '상한 = max(잠재량×20%, 기존)'],
    ['민감도 3', '민감도 2 + 송전 확장 × 기준배치', '민감도 2 + 송전 확장 × 재배치', '2035, 2040', '구간별 증설 최적화'],
], columns=['구분', '비교 기준', '비교 대상', '연도', '비고']), font=8, caption='표 0-1. 분석 구성')
H('입력 자료 점검 결과', 2)
B([('재배치 결과의 조건: ', '설비·송전망·비용이 상한 해제(OPEN) 결과와 완전히 같고, 반도체 수요 배분만 다름을 확인'),
   ('2035년 재배치 결과의 일반 전력수요: ', f"2035년 값({N('OPEN_2035', 'demand_EL_TWh'):.1f} TWh)이 아니라 "
    f"2040년 값({N('WP2_2035', 'demand_EL_TWh'):.1f} TWh)이 입력되어 있었음(11개 지역 모두 2040년 값과 동일). "
    '이를 그대로 비교하면 비용 감소 1.27조원이 재배치 효과로 오인되므로, 일반 수요만 2035년 값으로 되돌려 재최적화한 '
    '"보정" 결과로 비교함(저장된 기준 네트워크 재최적화 시 원래 결과와 0.0000% 차이로 재현됨을 확인)'),
   ('반도체 수요 총량: ', f"재배치 쪽이 2035년 +{N('WP2fix_2035', 'demand_Fab_TWh') - N('OPEN_2035', 'demand_Fab_TWh'):.1f} TWh, "
    f"2040년 +{N('WP2_2040', 'demand_Fab_TWh') - N('OPEN_2040', 'demand_Fab_TWh'):.1f} TWh 많음 → "
    f'수요 감축 분석의 MWh당 회피비용({AVOID:.0f} 천원)으로 비용 차이를 보정한 "재배치 순효과"를 함께 제시'),
   ('조치 권고: ', 'interface의 2035년 재배치 입력(일반 전력수요)을 확인 후 수정 필요')])
H('핵심 메시지', 2)
B([('1. 전국 비용과 전력믹스는 거의 같다. ', '재배치 순효과는 모든 조건에서 −0.22~−0.01조원/년(시스템 비용의 0.3% 이내). '
    'CO2 상한이 고정되어 원자력·가스·재생 발전량 구성은 변하지 않는다.'),
   ('2. 전력의 흐름이 크게 바뀐다. ', f"수도권 순수입이 2040년 {EX['OPEN_2040']['metro_import']:.0f} → "
    f"{EX['WP2_2040']['metro_import']:.0f} TWh로 절반 가까이 줄고, 송전 혼잡시간·HVDC 수송량·혼잡비용이 함께 감소한다."),
   ('3. 부담은 충청권으로 옮겨간다. ', '충북의 순수입이 크게 늘고, 상한이 없으면 충북 풍력이 잠재량을 초과(127%)하며, '
    f"REF 상한 조건에서는 공급부족이 대전세종충남에 집중된다(2040년 {r_('REF_2040', 'CND', 'slack_TWh'):.0f} → "
    f"{r_('REFWP2_2040', 'CND', 'slack_TWh'):.0f} TWh)."),
   ('4. 재배치는 송전 보강 수요를 줄인다. ', f"잠재량 상한 하에서 필요한 송전 증설이 2035년 {N('OPENcapgrid_2035', 'grid_added_GW'):.1f} → "
    f"{N('WP2capgrid_2035', 'grid_added_GW'):.1f} GW, 2040년 {N('OPENcapgrid_2040', 'grid_added_GW'):.1f} → "
    f"{N('WP2capgrid_2040', 'grid_added_GW'):.1f} GW로 감소하며, 특히 서울경기–충남 보강이 불필요해진다."),
   ('5. REF 상한의 공급부족은 배치로 해결되지 않는다. ', f"slack 총량은 2035년 {N('REF_2035', 'slack_TWh'):.0f} → "
    f"{N('REFWP2_2035', 'slack_TWh'):.0f} TWh, 2040년 {N('REF_2040', 'slack_TWh'):.0f} → {N('REFWP2_2040', 'slack_TWh'):.0f} TWh로 "
    '거의 같고(2040년 증가분은 반도체 수요 +2 TWh), 위치만 이동한다.')])

# 1장 재배치 내용
doc.add_page_break()
H('1. 반도체 수요 재배치 내용', 1)
fa = pd.DataFrame({'지역': [REGION_KR[r] for r in REG]})
for y in YEARS:
    fa[f'{y} 기준'] = R(f'OPEN_{y}').Fab_TWh.values
    fa[f'{y} 재배치'] = R(WP2[y]).Fab_TWh.values
    fa[f'{y} 차이'] = fa[f'{y} 재배치'] - fa[f'{y} 기준']
fa = fa[(fa.iloc[:, 1:].abs() > 0.01).any(axis=1)]
fa.loc[len(fa)] = ['합계'] + list(fa.iloc[:, 1:].sum().values)
T(fa, font=8, caption='표 1. 지역별 반도체 수요 (TWh)')
F('fig1_fab_allocation.png', '그림 1. 지역별 반도체 수요 배분 — 정부 발표 기준 vs 기업 설문 기반 재배치')
B([('수도권 → 충청권 이동: ', '서울경기의 반도체 수요가 전량 빠지고(2040년 −97 TWh), 충북(+51), 대전세종충남(+24), '
    '강원(+16), 부산울산·대구경북·경남(각 +4~7)으로 분산. 광주전남은 −8 TWh'),
   ('수도권 수요 비중: ', f"전력 버스 수요 중 수도권 비중이 2040년 {100 * R('OPEN_2040').loc[['GGD', 'ICN'], 'load_TWh'].sum() / R('OPEN_2040').load_TWh.sum():.0f}% → "
    f"{100 * R('WP2_2040').loc[['GGD', 'ICN'], 'load_TWh'].sum() / R('WP2_2040').load_TWh.sum():.0f}%로 감소")])

# 2장 상한 해제 조건 비교
doc.add_page_break()
H('2. 상한 해제 조건: 기준배치 vs 재배치', 1)
cols2 = [('OPEN_2035', '2035 기준'), ('WP2fix_2035', '2035 재배치'), ('OPEN_2040', '2040 기준'), ('WP2_2040', '2040 재배치')]
T(ind_table(cols2), font=7.5, caption='표 2. 주요 지표 비교 (2035 재배치는 일반수요 보정 결과)')
B([('비용: ', f"2035년 {s35.at['상한 해제', '비용 변화 (조원)']:+.2f}조원, 2040년 {s40.at['상한 해제', '비용 변화 (조원)']:+.2f}조원. "
    f"2040년은 반도체 수요가 2.0 TWh 많아 순효과는 {s40.at['상한 해제', '재배치 순효과 (조원)']:+.2f}조원으로 사실상 0")])
F('fig2_regional_change.png', '그림 2. 재배치에 따른 지역별 재생설비(a, b)와 순수입량(c) 변화')
H('주요 결과 및 해석', 2)
B([('재생에너지가 수요를 따라 이동: ', f"서울경기 풍력 2035년 {r_('WP2fix_2035', 'GGD', 'WT_GW') - r_('OPEN_2035', 'GGD', 'WT_GW'):+.1f} GW, "
    f"2040년 {r_('WP2_2040', 'GGD', 'WT_GW') - r_('OPEN_2040', 'GGD', 'WT_GW'):+.1f} GW 감소분을 충북 풍력"
    f"({r_('WP2fix_2035', 'CBD', 'WT_GW') - r_('OPEN_2035', 'CBD', 'WT_GW'):+.1f}, {r_('WP2_2040', 'CBD', 'WT_GW') - r_('OPEN_2040', 'CBD', 'WT_GW'):+.1f} GW)과 "
    '대전세종충남·부산울산 PV가 대체. 전국 PV·풍력 총량, 원자력·가스 발전량, ESS, 출력제한률은 거의 불변'),
   ('수도권 수입 의존 완화: ', f"수도권 순수입 2035년 {EX['OPEN_2035']['metro_import']:.0f} → {EX['WP2fix_2035']['metro_import']:.0f} TWh, "
    f"2040년 {EX['OPEN_2040']['metro_import']:.0f} → {EX['WP2_2040']['metro_import']:.0f} TWh. HVDC 수송량 2040년 "
    f"{EX['OPEN_2040']['hvdc']:.0f} → {EX['WP2_2040']['hvdc']:.0f} TWh"),
   ('송전 부담 감소(2040): ', f"수송거리 {N('OPEN_2040', 'grid_TWhkm'):,.0f} → {N('WP2_2040', 'grid_TWhkm'):,.0f} TWh·km, "
    f"혼잡시간 {N('OPEN_2040', 'grid_cong_h_sum'):,.0f} → {N('WP2_2040', 'grid_cong_h_sum'):,.0f} h, "
    f"혼잡비용 {N('OPEN_2040', 'grid_rent_조'):.2f} → {N('WP2_2040', 'grid_rent_조'):.2f}조원. "
    '2035년은 서울경기에 남은 풍력이 충남으로 역송되면서 혼잡시간이 소폭 증가'),
   ('충북 풍력의 잠재량 초과: ', f"충북 풍력이 기술적 잠재량의 {EX['OPEN_2040']['cbd_wt_pot']:.0f}% → {EX['WP2_2040']['cbd_wt_pot']:.0f}%(2040). "
    '대구경북→충북 구간이 기존부터 포화(혼잡 3,700~4,100시간)되어 있어 충북은 현지 풍력으로 수요를 충당 → 현실 입지 제약 반영 필요(민감도 2)'),
   ('반도체 입장의 지표: ', f"반도체 가중 전력가격 2040년 {N('OPEN_2040', 'price_Fab_loadw'):.1f} → {N('WP2_2040', 'price_Fab_loadw'):.1f} 천원/MWh로 동일, "
    f"반도체 입지 재생 자급률 {EX['OPEN_2040']['fab_re_ss']:.0f} → {EX['WP2_2040']['fab_re_ss']:.0f}%로 상승(RE100 조달 여건 개선)")])
F('fig3_grid_2040.png', '그림 3. 2040년 구간별 송전 수송량(a)과 혼잡시간(b)')

# 3장 REF vs OPEN
doc.add_page_break()
H('3. 재생 상한 설정(REF) vs 해제 (기준배치)', 1)
cols3 = [('REF_2035', '2035 REF'), ('OPEN_2035', '2035 상한해제'), ('REF_2040', '2040 REF'), ('OPEN_2040', '2040 상한해제')]
T(ind_table(cols3), font=7.5, caption='표 3. 상한 설정/해제 비교 (REF의 비용·가격은 slack 페널티가 결정하므로 직접 비교 불가)')
F('fig4_ref_vs_open_2040.png', '그림 4. 2040년 지역별 재생설비(a)와 REF의 지역별 공급부족(b)')
B([('공급부족: ', f"REF는 2035년 {N('REF_2035', 'slack_TWh'):.0f} TWh, 2040년 {N('REF_2040', 'slack_TWh'):.0f} TWh(전력 버스 수요의 "
    f"{100 * N('REF_2040', 'slack_TWh') / N('REF_2040', 'demand_ELbus_TWh'):.0f}%)의 slack 발생. 모든 지역 PV·풍력이 상한에 도달"),
   ('부족 지역: ', '서울경기·인천·충북·대전세종충남 — 반도체 수요가 몰린 지역과 일치'),
   ('상한 해제 시: ', f"PV+풍력 {N('REF_2040', 'PV_GW') + N('REF_2040', 'WT_GW'):.0f} → {N('OPEN_2040', 'PV_GW') + N('OPEN_2040', 'WT_GW'):.0f} GW(2040), "
    f"출력제한률 {N('REF_2040', 'curtail_%'):.1f} → {N('OPEN_2040', 'curtail_%'):.1f}%, 인천 PV가 잠재량의 139%까지 설치되는 등 입지 현실성 문제 발생"),
   ('해석: ', 'REF의 CO2 그림자가격(약 13,000 천원/t)은 실제 감축비용이 아니라 "상한 내에서 감축 목표 달성 불가"를 뜻하는 신호. '
    '송전망 민감도 분석에서 확인했듯 송전 확장으로는 해소되지 않는 전국적 공급 부족')])

# 4장 민감도 1
doc.add_page_break()
H('4. 민감도 1: REF 상한 조건에서의 재배치', 1)
H('분석 설정', 2)
B(['재배치 네트워크(2035는 일반수요 보정본)에 REF 결과의 PV·풍력 상한(p_nom_max)을 그대로 적용하여 재최적화',
   '비교: REF 기준배치 결과(기존)'])
cols4 = [('REF_2035', '2035 기준배치'), ('REFWP2_2035', '2035 재배치'), ('REF_2040', '2040 기준배치'), ('REFWP2_2040', '2040 재배치')]
T(ind_table(cols4).iloc[[0, 4, 5, 8, 9, 10, 12, 13]], font=7.5, caption='표 4. REF 상한 조건 비교')
F('fig5_ref_wp2_slack.png', '그림 5. REF 상한 조건의 지역별 공급부족(slack) — 기준배치 vs 재배치')
H('주요 결과 및 해석', 2)
B([('총량 불변: ', f"slack 2035년 {N('REF_2035', 'slack_TWh'):.1f} → {N('REFWP2_2035', 'slack_TWh'):.1f} TWh, "
    f"2040년 {N('REF_2040', 'slack_TWh'):.1f} → {N('REFWP2_2040', 'slack_TWh'):.1f} TWh(증가분은 반도체 수요 +2.0 TWh와 같음)"),
   ('위치 이동: ', f"서울경기 {r_('REF_2040', 'GGD', 'slack_TWh'):.0f} → {r_('REFWP2_2040', 'GGD', 'slack_TWh'):.0f}, "
    f"인천 {r_('REF_2040', 'ICN', 'slack_TWh'):.0f} → {r_('REFWP2_2040', 'ICN', 'slack_TWh'):.0f} TWh로 줄고, "
    f"대전세종충남이 {r_('REF_2040', 'CND', 'slack_TWh'):.0f} → {r_('REFWP2_2040', 'CND', 'slack_TWh'):.0f} TWh(지역 수요의 "
    f"{100 * r_('REFWP2_2040', 'CND', 'slack_TWh') / r_('REFWP2_2040', 'CND', 'load_TWh'):.0f}%)로 급증(2040). 2035년도 같은 양상"),
   ('송전: ', f"충북 순수입 {-r_('REF_2040', 'CBD', 'net_export_TWh'):.0f} → {-r_('REFWP2_2040', 'CBD', 'net_export_TWh'):.0f} TWh로 증가하고, "
    f"서울경기 순수입은 {-r_('REF_2040', 'GGD', 'net_export_TWh'):.0f} → {-r_('REFWP2_2040', 'GGD', 'net_export_TWh'):.0f} TWh로 감소. "
    '충북은 주변 송전망으로 전력을 끌어오고, 충남이 부족분을 떠안는 구조'),
   ('해석 주의: ', 'slack은 모든 지역에서 같은 비용이므로 위치는 송전 제약에 따라 정해지는 "부족의 출구"이다. 다만 반도체 수요가 '
    '충청권에 몰리면 이 지역의 공급 신뢰도 문제가 먼저 드러난다는 점은 정책적으로 의미가 있음')])
H('시사점', 2)
B(['REF 상한 하의 공급부족은 수요 배치로 해결되지 않으며, 재생 보급 상한 자체의 상향 또는 무탄소 공급원 확보가 선결 과제',
   '재배치 시 대전세종충남·충북의 공급 여력(재생 입지, 송전 접속)이 핵심 병목 → 충청권 공급 계획을 별도로 점검할 필요'])

# 5장 민감도 2
doc.add_page_break()
H('5. 민감도 2: 지역 잠재량 상한(20%) 하의 재배치', 1)
H('분석 설정', 2)
B(['지역별 PV·풍력 상한 = max(기술적 잠재량 × 20%, 기존 설비) — 이전 민감도 분석(세트 6)과 같은 방식',
   '상한 해제 조건의 "충북 풍력 잠재량 초과", "인천 PV 잠재량 초과" 문제를 제거한 상태에서 기준배치·재배치 비교'])
cols5 = [('OPENcap_2035', '2035 기준'), ('WP2cap_2035', '2035 재배치'), ('OPENcap_2040', '2040 기준'), ('WP2cap_2040', '2040 재배치')]
T(ind_table(cols5), font=7.5, caption='표 5. 잠재량 20% 상한 조건 비교')
F('fig6_cap_regional.png', '그림 6. 잠재량 상한 하에서 재배치에 따른 지역별 재생설비(a, b)와 순수입량(c) 변화')
H('주요 결과 및 해석', 2)
B([('상한 자체의 비용: ', f"상한 해제 대비 시스템 비용 2035년 +{N('OPENcap_2035', 'cost_excl_penalty_조') - N('OPEN_2035', 'cost_excl_penalty_조'):.2f}, "
    f"2040년 +{N('OPENcap_2040', 'cost_excl_penalty_조') - N('OPEN_2040', 'cost_excl_penalty_조'):.2f}조원/년(기준배치)"),
   ('재배치 순효과: ', f"2035년 {s35.at['잠재량 20%', '재배치 순효과 (조원)']:+.2f}조원(재배치가 유리), 2040년 "
    f"{s40.at['잠재량 20%', '재배치 순효과 (조원)']:+.2f}조원. 상한이 있을 때 재배치의 이점이 조금 더 커짐(2035)"),
   ('재생 입지 이동: ', f"2035년 강원 PV {r_('WP2cap_2035', 'GWD', 'PV_GW') - r_('OPENcap_2035', 'GWD', 'PV_GW'):+.1f} GW, 대전세종충남 PV "
    f"{r_('WP2cap_2035', 'CND', 'PV_GW') - r_('OPENcap_2035', 'CND', 'PV_GW'):+.1f} GW로 반도체 수요가 늘어난 충청권 인근으로 PV가 이동. 2040년은 대부분 지역이 상한에 도달해 변화가 1~2 GW 이내"),
   ('충북은 수입으로 충당: ', f"충북 풍력이 상한(5 GW)에 묶여 순수입이 2040년 {-r_('OPENcap_2040', 'CBD', 'net_export_TWh'):.0f} → "
    f"{-r_('WP2cap_2040', 'CBD', 'net_export_TWh'):.0f} TWh로 급증. 충남→충북 수송량 증가, 대구경북→충북 혼잡 "
    f"{corr[(corr.scenario == 'OPENcap_2040') & (corr.corridor == 'GBD-CBD')].cong_h.iloc[0]:,.0f} → "
    f"{corr[(corr.scenario == 'WP2cap_2040') & (corr.corridor == 'GBD-CBD')].cong_h.iloc[0]:,.0f}시간"),
   ('수도권 의존 완화는 유지: ', f"수도권 순수입 2040년 {EX['OPENcap_2040']['metro_import']:.0f} → {EX['WP2cap_2040']['metro_import']:.0f} TWh, "
    f"혼잡시간 합계 {N('OPENcap_2040', 'grid_cong_h_sum'):,.0f} → {N('WP2cap_2040', 'grid_cong_h_sum'):,.0f} h. 서울경기–충남 혼잡이 "
    f"{corr[(corr.scenario == 'OPENcap_2040') & (corr.corridor == 'GGD-CND')].cong_h.iloc[0]:,.0f} → "
    f"{corr[(corr.scenario == 'WP2cap_2040') & (corr.corridor == 'GGD-CND')].cong_h.iloc[0]:,.0f}시간으로 크게 감소"),
   ('반도체 부담 가격: ', f"반도체 가중 전력가격 2035년 {N('OPENcap_2035', 'price_Fab_loadw'):.1f} → {N('WP2cap_2035', 'price_Fab_loadw'):.1f}, "
    f"2040년 {N('OPENcap_2040', 'price_Fab_loadw'):.1f} → {N('WP2cap_2040', 'price_Fab_loadw'):.1f} 천원/MWh로 소폭 하락")])
H('시사점', 2)
B(['현실적 입지 제약을 두면 재배치는 "충북 현지 재생"이 아니라 "충청권으로의 송전"으로 충당되므로, 충북 접속 송전망이 핵심',
   '상한 해제 결과(충북 풍력 127%)는 과대평가이며, 정책 보고에는 잠재량 상한을 둔 결과를 기준으로 제시하는 것이 바람직'])

# 6장 민감도 3
doc.add_page_break()
H('6. 민감도 3: 잠재량 상한 + 필요 구간 송전 확장', 1)
H('분석 설정', 2)
B(['민감도 2 조건에 모든 AC 구간(양방향 증설 링크)·HVDC 확장을 허용 → 최적화가 증설 구간·규모 결정(건설비 × 1, 기존의 3배까지)',
   '건설비: AC 345kV 2,000 천원/MW·km(최소 30 km), HVDC 3,000 천원/MW·km + 변환소 300,000 천원/MW, 연간화 계수 0.064'])
cols6 = [('OPENcapgrid_2035', '2035 기준'), ('WP2capgrid_2035', '2035 재배치'), ('OPENcapgrid_2040', '2040 기준'),
         ('WP2capgrid_2040', '2040 재배치')]
T(ind_table(cols6), font=7.5, caption='표 6. 잠재량 상한 + 송전 확장 조건 비교')
F('fig7_grid_expansion.png', '그림 7. 필요 구간 송전 증설량 — 기준배치 vs 재배치')
ex = corr[corr.scenario.isin([c for c, _ in cols6])].pivot_table(index=['type', 'corridor'], columns='scenario', values='added_GW')
ex = ex[ex.max(axis=1) > 0.1][[c for c, _ in cols6]].round(2)
ex.columns = [l for _, l in cols6]
ex = ex.reset_index()
ex['corridor'] = ex.corridor.map(kr_corr)
T(ex.rename(columns={'type': '유형', 'corridor': '구간'}), font=8, caption='표 7. 구간별 송전 증설량 (GW)')
H('주요 결과 및 해석', 2)
B([('증설 필요량 감소: ', f"2035년 {N('OPENcapgrid_2035', 'grid_added_GW'):.1f} → {N('WP2capgrid_2035', 'grid_added_GW'):.1f} GW, "
    f"2040년 {N('OPENcapgrid_2040', 'grid_added_GW'):.1f} → {N('WP2capgrid_2040', 'grid_added_GW'):.1f} GW. "
    f"연간 증설비 {N('OPENcapgrid_2040', 'grid_exp_cost_조'):.2f} → {N('WP2capgrid_2040', 'grid_exp_cost_조'):.2f}조원(2040)"),
   ('서울경기–충남 보강 불필요: ', '기준배치에서 필요한 서울경기–충남 증설(2035년 11.4 GW, 2040년 3.5 GW)이 재배치 시 5.3 GW, 0 GW로 감소. '
    '수도권 반도체 수요가 빠지면서 충청→수도권 송전 필요가 줄어든 결과'),
   ('공통 증설 구간: ', '강원–대구경북(2035년 약 19~20 GW, 2040년 약 12 GW)은 두 배치 모두 최대 증설 구간 → 동해안 발전·재생 '
    '전력의 남북 연계는 반도체 배치와 무관한 핵심 보강 구간. 2035년 대구경북–부산울산(약 8.8 GW), 경남–광주전남(6.6 GW)도 공통'),
   ('비용: ', f"송전 확장 후 시스템 비용은 2035년 두 배치가 {N('OPENcapgrid_2035', 'cost_excl_penalty_조'):.2f}조원으로 같고, 2040년 재배치 순효과는 "
    f"{s40.at['잠재량 20% +송전 확장', '재배치 순효과 (조원)']:+.2f}조원. 송전 확장의 비용 절감 효과는 기준배치 "
    f"{N('OPENcap_2040', 'cost_excl_penalty_조') - N('OPENcapgrid_2040', 'cost_excl_penalty_조'):.2f} vs 재배치 "
    f"{N('WP2cap_2040', 'cost_excl_penalty_조') - N('WP2capgrid_2040', 'cost_excl_penalty_조'):.2f}조원(2040)으로, "
    '재배치가 송전 보강의 일부를 대체'),
   ('대구경북–충북 미증설: ', '혼잡이 2,000시간 이상 남지만 증설되지 않음 → 이 구간 대신 다른 경로 보강과 충북 PV가 더 경제적이라는 의미')])
H('시사점', 2)
B(['반도체 재배치는 수도권 방향 송전 보강 부담을 줄이는 "수요 측 송전 대안"의 성격을 가짐',
   '강원–대구경북 등 동해안 연계 보강은 어느 배치에서도 필요한 "후회 없는(no-regret)" 투자 후보',
   '모형 한계: 신규 송전은 조류 제어 가능한 링크로 근사(KVL 미적용), 실제 경과지·단가 반영 시 재검토 필요'])

# 7장 종합
doc.add_page_break()
H('7. 종합 및 정책 시사점', 1)
F('fig8_summary.png', '그림 8. 조건별 재배치 순효과(a), 수도권 순수입 변화(b), 송전 혼잡시간 변화(c)')
sm = summ[['연도', '조건', '비용 변화 (조원)', '재배치 순효과 (조원)', '수도권 순수입 기준 (TWh)', '수도권 순수입 재배치 (TWh)',
           '혼잡시간 기준 (h)', '혼잡시간 재배치 (h)']].copy()
T(sm, font=7.5, caption=f'표 8. 조건별 재배치 효과 요약 (순효과 = 비용 변화 − 반도체 수요 차이 × {AVOID:.0f} 천원/MWh)')
H('종합 판단', 2)
B([('비용 중립적: ', '어느 조건에서도 재배치의 전국 비용 순효과는 −0.22~−0.01조원/년으로 미미 → 기업 설문 기반 배치를 채택해도 시스템 비용 부담은 크지 않음'),
   ('공간 구조의 전환: ', '"수도권 수요 – 지방 공급 – 장거리 송전" 구조에서 "충청권 수요 – 인근 공급" 구조로 이동. 수도권 순수입·혼잡·HVDC 의존이 줄어듦'),
   ('새로운 병목: ', '충북·대전세종충남의 공급 여력과 접속 송전망이 핵심. REF 상한에서는 충남 공급부족, 잠재량 상한에서는 충북 수입 급증으로 나타남'),
   ('송전 계획과의 연계: ', '재배치 시 서울경기–충남 보강 필요가 줄고, 강원–대구경북 등 동해안 연계는 공통 필요 구간'),
   ('REF 공급부족: ', '반도체 배치와 무관하게 재생 보급 상한이 공급부족의 근본 원인')])
H('새로 도입한 지표', 2)
T(pd.DataFrame([
    ['재배치 순효과', '비용 변화 − 수요 차이 × MWh당 회피비용', '수요 총량이 약간 다른 결과 간 순수 입지 효과 분리'],
    ['수도권 순수입', '서울경기·인천의 (유입 − 유출) 송전량', '수도권 전력 의존도, 수요 분산 효과'],
    ['송전 수송거리', 'Σ 구간 수송량 × 구간 거리 (TWh·km)', '장거리 송전 부담'],
    ['혼잡시간·혼잡비용', '용량 99% 이상 시간 / Σ 수송량 × 양단 가격차', '병목 위치와 경제적 손실'],
    ['반도체 가중 전력가격', 'Σ(반도체 수요 × 지역 한계가격) / Σ반도체 수요', '반도체 기업이 체감하는 전력 비용'],
    ['반도체 입지 재생 자급률', 'Σ(반도체 수요 × 지역 재생발전/지역수요) / Σ반도체 수요', 'RE100 조달 여건'],
    ['잠재량 대비 설치율', '지역 설비 / 기술적 잠재량', '결과의 입지 현실성 점검'],
], columns=['지표', '정의', '용도']), font=8, caption='표 9. 이번 분석에서 사용한 추가 지표')
H('추가 검토 과제', 2)
B(['interface의 2035년 재배치 입력(일반 전력수요) 수정 후 재실행',
   '반도체 RE100 조달 제약: 재생 전용 반도체 버스(DC_Fab_RE)로 반도체 수요의 30~50%를 이전해 입지별 RE 조달비용 비교 '
   '(현재 계통→전용 버스 링크 용량 0, 구조 보완 필요)',
   '재배치 상태에서 반도체 수요 규모 ±25% — 충청권 집중에 따른 지역 리스크',
   'REF 상한 + 재배치 + 송전 확장 조합, 충청권 재생 상한 완화 시나리오',
   '정부안과 설문안의 중간 배분(50:50)으로 효과의 선형성 점검'])

# 부록
doc.add_page_break()
H('부록. 방법 및 결과 파일', 1)
B([('비교 방식: ', '저장된 결과 네트워크(.nc)를 불러와 조건만 바꿔 재최적화(PyPSA_GUI.optimize_network 사용). '
    '기준 네트워크 재최적화 시 원래 결과와 목적함수 0.0000% 차이로 재현 확인'),
   ('시스템 비용: ', '목적함수 − slack·열 보완 발전 페널티(5,000 천원/MWh). 기존 설비 고정비는 제외(상수)'),
   ('MWh당 시스템 비용: ', '시스템 비용 ÷ 전체 최종수요(전력·EV·열 포함)'),
   ('전력가격: ', '전력 버스 한계가격(LMP)의 수요 가중 평균'),
   ('REF 결과 해석: ', 'slack이 있는 경우 한계가격·CO2 그림자가격·혼잡비용이 페널티에 의해 결정되므로 상한 해제 결과와 직접 비교 불가')])
T(pd.DataFrame([
    ['기존 결과(복사본)', r'wp2_compare_nc\{REF,OPEN,WP2}_{2035,2040}.nc', '한글 경로를 netCDF가 읽지 못해 복사'],
    ['2035 재배치 보정', r'wp2_compare_nc\WP2fix_2035.nc', '일반 수요를 2035년 값으로 복원'],
    ['민감도 1~3', r'wp2_compare_nc\{REFWP2,OPENcap,WP2cap,OPENcapgrid,WP2capgrid}_{연도}.nc', 'run_wp2_sensitivity.py'],
    ['지표 계산', rf'{D}\national.csv, regional.csv, corridor.csv', 'compare_wp2.py'],
    ['보고서·표·그림', rf'{D}\WP2_비교_민감도_보고서.docx, WP2_비교_민감도_표.xlsx', 'make_wp2_report.py'],
], columns=['내용', '파일', '비고']), font=7.5)

path = os.path.join(OUT, 'WP2_비교_민감도_보고서.docx')
doc.save(path)
print('완료:', path)
print(summ.round(2).to_string())
