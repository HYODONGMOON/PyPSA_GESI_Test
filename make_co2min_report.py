"""
비용 상한 없는 배출량 최소화 — 2035·2040 감축목표 도달 가능성 보고서 (Word + 그림 + 표)

선행:
  python run_co2_min.py ... (2035, 2040)
  python make_co2cost_report.py ... (비용 민감도 national.csv, 경로 그림용)
사용법:
  python make_co2min_report.py [--out reports/co2min_<날짜>]
"""
import argparse
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

from analyze_mga import analyze, kind
from run_co2_cost_sweep import aux_capex, aux_gens
from run_mga_sweep import penalty_value

logging.disable(logging.WARNING)
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 9

ap = argparse.ArgumentParser()
ap.add_argument('--out', default=f"reports/co2min_{datetime.now():%Y%m%d}")
args = ap.parse_args()
OUT = args.out
os.makedirs(OUT, exist_ok=True)

YEARS = {
    2035: dict(target=88.3,
               B0='results/co2_price_search_2035_20261110_065714/B0.nc',
               MIN='results/_tmp_ref/co2min2035/CO2min-lam1000.nc', lam_min=1000,
               checks=[('λ=500', 'results/co2_price_search_2035_20261110_065714/lambda/lam_500.nc', 500, '비용 민감도 탐색'),
                       ('λ=700', 'results/_tmp_ref/co2min2035/CO2min-lam700.nc', 700, ''),
                       ('λ=1,000', 'results/_tmp_ref/co2min2035/CO2min-lam1000.nc', 1000, '채택')],
               sens='reports/co2cost2035_20261110/national.csv',
               ref=r'results\대표경로\1105\3. 설비고정+CO2CAPX\2035'),
    2040: dict(target=58.9,
               B0='results/co2_cost_sweep_2040_20261109_145316/B0.nc',
               MIN='results/co2_min_2040_20261110_144137/CO2min-lam1000.nc', lam_min=1000,
               checks=[('λ=500', 'results/co2_price_search_2040_20261109_162249/lambda/lam_500.nc', 500, '비용 민감도 탐색'),
                       ('λ=1,000', 'results/co2_min_2040_20261110_144137/CO2min-lam1000.nc', 1000, '채택'),
                       ('λ=3,000', 'results/co2_min_2040_20261110_144137/CO2min-lam3000.nc', 3000, '제외 (보조공급 증가)'),
                       ('λ=3,000 + 보조공급 상한', 'results/_tmp_ref/co2min2040cap/CO2min-lam3000-auxcap.nc', 3000,
                        '검증 (보조공급 ≤ 기준+1GWh)'),
                       ('λ=10,000', 'results/co2_price_search_2040_20261109_162249/lambda/lam_10000.nc', 10000,
                        '제외 (보조공급 증가)')],
               sens='reports/co2cost2040_20261110/national.csv',
               ref=r'results\대표경로\1105\3. 설비고정+CO2CAPX\2040'),
}
C_COAL, C_LNG, C_CHP, C_NU, C_PV, C_WT, C_FC = '#5D4037', '#E07B39', '#8E5EA2', '#6C757D', '#F2B134', '#3C8DBC', '#2E7D32'


def metrics(n):
    o, _ = analyze(n, '')
    w = n.snapshot_weightings.generators
    g, em = n.generators, n.carriers['co2_emissions']
    eff = n.get_switchable_as_dense('Generator', 'efficiency')
    P = n.generators_t.p
    coal = g.index[g.carrier == 'coal']
    lng = [x for x in g.index[g.carrier == 'gas'] if x.endswith('_LNG')]
    fuel = [x for x in g.index[g.carrier == 'gas'] if 'Fuel_Supply' in x]
    for key, idx, car in (('coal', coal, 'coal'), ('LNG', lng, 'gas'), ('CHPfuel', fuel, 'gas')):
        o[f'CO2_{key}_Mt'] = float((P[idx] / eff[idx]).mul(w, axis=0).sum().sum() * em[car] / 1e6)
    o['coal_TWh'] = float(P[coal].mul(w, axis=0).sum().sum() / 1e6)
    o['coal_GW'] = g.loc[coal, 'p_nom'].sum() / 1e3
    o['LNG_GW'] = g.loc[lng, 'p_nom'].sum() / 1e3
    o['nuclear_GW'] = g.loc[g.index[g.carrier == 'nuclear'], 'p_nom'].sum() / 1e3
    o['coal_CF_%'] = 100 * o['coal_TWh'] * 1e3 / (o['coal_GW'] * w.sum()) if o['coal_GW'] else np.nan
    o['LNG_CF_%'] = 100 * o['LNG_power_TWh'] * 1e3 / (o['LNG_GW'] * w.sum())
    o['nuclear_CF_%'] = 100 * o['nuclear_TWh'] * 1e3 / (o['nuclear_GW'] * w.sum())
    o['LNG_tCO2_per_MWh'] = o['CO2_LNG_Mt'] / o['LNG_power_TWh'] if o['LNG_power_TWh'] else np.nan
    o['aux_TWh'] = float(P[aux_gens(n)].mul(w, axis=0).sum().sum() / 1e6)
    o['power_slack_TWh'] = float(P[[x for x in g.index if 'slack' in x.lower() or 'failsafe' in x.lower()]]
                                 .mul(w, axis=0).sum().sum() / 1e6)
    o['Z'] = float(n.objective)
    o['C_sys'] = o['Z'] - penalty_value(n) - aux_capex(n)
    o['penalty'] = penalty_value(n)
    o['slack_capex'] = aux_capex(n)
    caps = {}
    for comp, attr in (('Generator', 'p_nom'), ('Link', 'p_nom'), ('Store', 'e_nom')):
        df = n.static(comp)
        for nm in df.index[df[f'{attr}_extendable'].astype(bool)]:
            k = {'solar': 'PV', 'wind': 'WT'}.get(df.at[nm, 'carrier'], kind(nm)) if comp == 'Generator' else kind(nm)
            u = 'GWh' if comp == 'Store' else 'GW'
            caps[(k, u)] = caps.get((k, u), 0.0) + float(df.at[nm, f'{attr}_opt']) / 1e3
            caps[(k + '|exist', u)] = caps.get((k + '|exist', u), 0.0) + float(df.at[nm, attr]) / 1e3
    return o, caps


R, CAPS, CHK, SENS = {}, {}, {}, {}
for y, cfg in YEARS.items():
    for c in ('B0', 'MIN'):
        print(f'[분석] {y} {c}', flush=True)
        n = pypsa.Network(cfg[c])
        R[(y, c)], CAPS[(y, c)] = metrics(n)
        if c == 'MIN':
            R[(y, c)]['Z'] = float(n.objective)
        del n
    b = R[(y, 'B0')]
    rows = []
    for lbl, p, lam, note in cfg['checks']:
        n = pypsa.Network(p)
        o, _ = metrics(n)
        rows.append({'경우': lbl, '탄소가격 λ (천원/t)': lam, 'CO2 (Mt)': o['CO2_Mt'],
                     '시스템 비용 증가 (%, slack 제외)': 100 * (o['C_sys'] / b['C_sys'] - 1),
                     '보조공급 변화 (GWh)': 1e3 * (o['aux_TWh'] - b['aux_TWh']), '비고': note})
        del n
    CHK[y] = pd.DataFrame(rows)
    s = pd.read_csv(cfg['sens'], index_col=0)
    SENS[y] = s[['cost_increase_pct', 'CO2_Mt']]
    m_ = R[(y, 'MIN')]
    m_['gap_Mt'] = m_['CO2_Mt'] - cfg['target']
    m_['need_TWh'] = m_['gap_Mt'] / m_['LNG_tCO2_per_MWh']
    for c in ('B0', 'MIN'):
        r = R[(y, c)]
        r['dC_jo'] = (r['C_sys'] - b['C_sys']) / 1e9
        r['dC_%'] = 100 * (r['C_sys'] / b['C_sys'] - 1)
        r['dZ_%'] = 100 * (r['Z'] / b['Z'] - 1)
        r['dE_Mt'] = r['CO2_Mt'] - b['CO2_Mt']
    m_['avg_abate_kwon_t'] = 1e3 * m_['dC_jo'] / -m_['dE_Mt']


def V(y, c, k):
    return float(R[(y, c)][k])


figs = {}


def save(fig, name):
    p = os.path.join(OUT, name)
    fig.tight_layout()
    fig.savefig(p, dpi=200)
    plt.close(fig)
    figs[name] = p


# 그림 1: 비용 증가 vs 최소 배출 경로 + 감축목표
fig, ax = plt.subplots(1, 2, figsize=(11.5, 4.2))
for a, (y, cfg) in zip(ax, YEARS.items()):
    s = SENS[y]
    a.plot(s.cost_increase_pct, s.CO2_Mt, '-o', color='#2E75B6', ms=4, label='비용 민감도 (+0~10%)')
    xm, ym = V(y, 'MIN', 'dC_%'), V(y, 'MIN', 'CO2_Mt')
    a.plot([s.cost_increase_pct.iloc[-1], xm], [s.CO2_Mt.iloc[-1], ym], ':', color='#2E75B6')
    a.plot(xm, ym, '*', color='#C0392B', ms=14, label=f'배출 최소 ({ym:.1f} Mt)')
    a.axhline(cfg['target'], color='#2E7D32', ls='--', lw=1.2)
    a.text(1, cfg['target'] + 2, f"감축목표 {cfg['target']} Mt", color='#2E7D32', fontsize=8.5)
    a.annotate('', xy=(xm, cfg['target']), xytext=(xm, ym), arrowprops=dict(arrowstyle='<->', color='#C0392B'))
    a.text(xm - 1, (ym + cfg['target']) / 2, f"부족 {ym - cfg['target']:.1f} Mt", color='#C0392B', ha='right', fontsize=8.5)
    a.set_title(f'({"a" if y == 2035 else "b"}) {y}년')
    a.set_xlabel('기준 대비 시스템 비용 증가율 (%, slack 비용 제외)'); a.set_ylabel('CO2 배출량 (Mt)')
    a.set_ylim(min(cfg['target'] - 15, 0) if cfg['target'] < 15 else cfg['target'] - 15, V(y, 'B0', 'CO2_Mt') + 15)
    a.grid(alpha=0.3); a.legend(fontsize=8, loc='upper right')
save(fig, 'fig1_path_target.png')

# 그림 2: 배출원별 CO2 (기준 vs 최소) + 목표
fig, ax = plt.subplots(figsize=(8.5, 4))
labels, xs = [], []
for i, (y, cfg) in enumerate(YEARS.items()):
    for j, c in enumerate(('B0', 'MIN')):
        x = i * 2.6 + j
        bot = 0
        for key, nm, col in (('CO2_coal_Mt', '석탄', C_COAL), ('CO2_LNG_Mt', 'LNG 발전', C_LNG), ('CO2_CHPfuel_Mt', 'CHP 연료', C_CHP)):
            v = V(y, c, key)
            ax.bar(x, v, 0.75, bottom=bot, color=col, label=nm if (i, j) == (0, 0) else None)
            bot += v
        ax.text(x, bot + 3, f'{bot:.1f}', ha='center', fontsize=8.5)
        xs.append(x); labels.append(f"{y}\n{'기준' if c == 'B0' else '배출 최소'}")
    ax.hlines(cfg['target'], i * 2.6 - 0.5, i * 2.6 + 1.5, color='#2E7D32', ls='--', lw=1.4,
              label='감축목표' if i == 0 else None)
    ax.text(i * 2.6 + 0.5, cfg['target'] + 3, f"목표 {cfg['target']}", color='#2E7D32', fontsize=8, ha='center',
            bbox=dict(facecolor='white', edgecolor='none', pad=1))
ax.set_xticks(xs, labels); ax.set_ylabel('Mt'); ax.legend(fontsize=8, ncol=4, loc='upper right')
ax.set_ylim(0, max(V(y, 'B0', 'CO2_Mt') for y in YEARS) * 1.18); ax.grid(axis='y', alpha=0.3)
ax.set_title('배출원별 CO2: 기준 vs 배출 최소')
save(fig, 'fig2_co2_source.png')

# 그림 3: 발전 믹스 + 목표 도달에 필요한 추가 무탄소 전력
mix = [('석탄', 'coal_TWh', C_COAL), ('LNG 발전', 'LNG_power_TWh', C_LNG), ('CHP 발전', 'CHP_power_TWh', C_CHP),
       ('원자력', 'nuclear_TWh', C_NU), ('PV', 'PV_TWh', C_PV), ('풍력', 'WT_TWh', C_WT), ('연료전지', 'fuelcell_power_TWh', C_FC)]
fig, ax = plt.subplots(figsize=(10, 4.2))
xs, labels = [], []
for i, y in enumerate(YEARS):
    for j, c in enumerate(('B0', 'MIN')):
        x = i * 2.6 + j
        bot = 0
        for nm, k, col in mix:
            v = V(y, c, k)
            ax.bar(x, v, 0.75, bottom=bot, color=col, label=nm if (i, j) == (0, 0) else None)
            bot += v
        if c == 'MIN':
            need = V(y, 'MIN', 'need_TWh')
            lngv = V(y, 'MIN', 'LNG_power_TWh')
            ax.bar(x + 0.42, need, 0.12, bottom=V(y, 'MIN', 'coal_TWh') + lngv - need, color='none', edgecolor='#2E7D32',
                   hatch='///', label='목표 도달에 필요한\nLNG → 무탄소 대체' if i == 0 else None)
        xs.append(x); labels.append(f"{y}\n{'기준' if c == 'B0' else '배출 최소'}")
ax.set_xticks(xs, labels); ax.set_ylabel('TWh'); ax.legend(fontsize=7.5, loc='center left', bbox_to_anchor=(1.01, 0.5))
ax.grid(axis='y', alpha=0.3); ax.set_title('발전원별 발전량')
save(fig, 'fig3_mix.png')


# ───────────── 표 ─────────────
def summary_table():
    rows = []
    for y, cfg in YEARS.items():
        for c in ('B0', 'MIN'):
            rows.append({'연도': y, '경우': '기준 (비용최소)' if c == 'B0' else '배출 최소',
                         'CO2 (Mt)': V(y, c, 'CO2_Mt'), '감축목표 (Mt)': cfg['target'],
                         '목표 대비 (Mt)': V(y, c, 'CO2_Mt') - cfg['target'],
                         '시스템 비용 (조원, slack 제외)': V(y, c, 'C_sys') / 1e9, '비용 증가 (조원)': V(y, c, 'dC_jo'),
                         '증가율 (%, slack 제외)': V(y, c, 'dC_%'), '증가율 (%, 목적함수)': V(y, c, 'dZ_%'),
                         '평균 감축비용 (천원/t)': np.nan if c == 'B0' else V(y, c, 'avg_abate_kwon_t')})
    return pd.DataFrame(rows)


def balance_table():
    its = [('석탄 발전', 'coal_TWh'), ('LNG 발전', 'LNG_power_TWh'), ('CHP 발전', 'CHP_power_TWh'), ('CHP 열', 'CHP_heat_TWh'),
           ('원자력', 'nuclear_TWh'), ('PV 발전', 'PV_TWh'), ('풍력 발전', 'WT_TWh'), ('출력제한', 'curtail_TWh'),
           ('연료전지 발전', 'fuelcell_power_TWh'), ('히트펌프 전력', 'HP_elec_TWh'), ('히트펌프 열', 'HP_heat_TWh'),
           ('수전해 전력', 'electrolyser_elec_TWh'), ('ESS 충전', 'ESS_charge_TWh'), ('V2G 방전', 'V2G_TWh'),
           ('열보조공급', 'heat_fallback_TWh'), ('전력 slack', 'power_slack_TWh'), ('HVDC 수송량', 'HVDC_TWh'),
           ('석탄 이용률 (%)', 'coal_CF_%'), ('LNG 이용률 (%)', 'LNG_CF_%'), ('원자력 이용률 (%)', 'nuclear_CF_%'),
           ('CO2 석탄 (Mt)', 'CO2_coal_Mt'), ('CO2 LNG 발전 (Mt)', 'CO2_LNG_Mt'), ('CO2 CHP 연료 (Mt)', 'CO2_CHPfuel_Mt'),
           ('CO2 합계 (Mt)', 'CO2_Mt')]
    df = pd.DataFrame({'항목 (TWh)': [a for a, _ in its]})
    for y in YEARS:
        for c in ('B0', 'MIN'):
            df[f"{y} {'기준' if c == 'B0' else '최소'}"] = [V(y, c, k) for _, k in its]
    return df


GROUPS = [('석탄 연료비', ['c_opex_coal']), ('가스 연료비', ['c_opex_gas']), ('CHP 운전비', ['c_opex_CHP']),
          ('원자력 연료비', ['c_opex_nuclear'])]


def cost_table():
    rows = [g for g, _ in GROUPS] + ['기타 (투자비·기타 운전비)', '시스템 비용 합계', '(제외) slack·Fallback 운전 페널티',
                                     '(제외) slack 설비 투자비']
    df = pd.DataFrame({'비용 항목 (조원/년)': rows})
    for y in YEARS:
        for c in ('B0', 'MIN'):
            r = R[(y, c)]
            vals = [sum(float(r.get(k, 0.0)) for k in ks) for _, ks in GROUPS]
            tot = r['C_sys'] / 1e9
            df[f"{y} {'기준' if c == 'B0' else '최소'}"] = vals + [tot - sum(vals), tot, r['penalty'] / 1e9,
                                                                  r['slack_capex'] / 1e9]
    return df


KR = {'PV': '태양광', 'WT': '풍력', 'ESS': 'ESS 저장', 'ESS_Ch': 'ESS 충전', 'ESS_Disch': 'ESS 방전', 'Electrolyser': '수전해',
      'Fuelcell': '연료전지', 'H2_Storage': '수소 저장', 'H_Storage': '열 저장', 'EV_DSM': 'EV DSM', 'slack': '전력 slack',
      'H_Fallback_Gen': '열보조공급', 'LNG_Fuel_Supply': 'CHP 연료공급'}


def capacity_table():
    rows = []
    for y in YEARS:
        cb, cm = CAPS[(y, 'B0')], CAPS[(y, 'MIN')]
        for (k, u) in sorted(x for x in cb if '|exist' not in x[0]):
            if abs(cm[(k, u)] - cb[(k, u)]) > 0.01 or k in ('PV', 'WT'):
                rows.append({'연도': y, '설비': f'{KR.get(k, k)} ({u})', '입력 기존': cb[(k + '|exist', u)],
                             '기준': cb[(k, u)], '배출 최소': cm[(k, u)], '변화': cm[(k, u)] - cb[(k, u)]})
    return pd.DataFrame(rows)


TBL = {'표1_요약': summary_table(), '표2_수급': balance_table(), '표3_비용': cost_table(), '표4_설비': capacity_table()}
for y in YEARS:
    TBL[f'표5_검증_{y}'] = CHK[y]
with pd.ExcelWriter(os.path.join(OUT, 'CO2_최소배출_2035_2040_표.xlsx')) as xw:
    for k, v in TBL.items():
        v.to_excel(xw, sheet_name=k, index=False)
    pd.DataFrame({f'{y} {c}': R[(y, c)] for y in YEARS for c in ('B0', 'MIN')}).drop(index=['case'], errors='ignore') \
        .to_excel(xw, sheet_name='전체지표')

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


def fy(y, k, f='{:.1f}'):
    return f.format(V(y, 'MIN', k))


doc.add_heading('2035·2040년 배출량 최소화: 감축목표 도달 가능성 분석', 0)
P('설비 고정 + CO2 배출 상한 해제 조건에서 비용 상한 없이 배출량을 최소화했을 때의 최소 배출량과 필요 비용', size=12)
P(f'작성일: {datetime.now():%Y-%m-%d}   |   모형: PyPSA 0.31.1 + linopy + CPLEX Barrier, 8,784시간 단일연도', size=9)

H = doc.add_heading
H('개요', 1)
P('비용 민감도 분석(기준 비용 대비 +0~10%)에 이어, 비용 상한을 완전히 없애고 배출량을 최소화하였다. 목적은 2035년·2040년 감축목표'
  '(각 88.3 Mt, 58.9 Mt)에 비용을 얼마나 더 쓰면 도달할 수 있는지, 혹은 현재 설비로는 도달할 수 없는지를 확인하는 것이다. '
  '입력은 두 연도 모두 비용 민감도 분석과 같은 조건(설비 고정, CO2 상한 비구속)이다.')

H('핵심 결과', 2)
B([('결론: ', '비용 상한을 없애도 두 연도 모두 감축목표에 도달할 수 없다. 현재 설비 구성으로는 배출을 '
    f"2035년 {fy(2035, 'CO2_Mt')} Mt, 2040년 {fy(2040, 'CO2_Mt')} Mt까지만 줄일 수 있어, 목표보다 각각 "
    f"{fy(2035, 'gap_Mt')} Mt, {fy(2040, 'gap_Mt')} Mt 많다."),
   ('필요 비용: ', f"최소 배출 운영의 시스템 비용(slack 제외)은 2035년 {V(2035, 'MIN', 'C_sys') / 1e9:.2f}조원"
    f"(기준 대비 +{V(2035, 'MIN', 'dC_jo'):.2f}조원, +{V(2035, 'MIN', 'dC_%'):.1f}%), 2040년 {V(2040, 'MIN', 'C_sys') / 1e9:.2f}조원"
    f"(+{V(2040, 'MIN', 'dC_jo'):.2f}조원, +{V(2040, 'MIN', 'dC_%'):.1f}%). 목적함수 기준 증가율은 각각 "
    f"+{V(2035, 'MIN', 'dZ_%'):.1f}%, +{V(2040, 'MIN', 'dZ_%'):.1f}%."),
   ('원인: ', f"최소 배출 운영에서 석탄은 거의 쓰지 않는다(2035년 {fy(2035, 'coal_TWh')} TWh, 2040년 {fy(2040, 'coal_TWh')} TWh). "
    '남은 배출은 전부 가스(LNG 발전 + CHP 연료)다. 원자력은 이미 최대로 운전하고 재생에너지 용량은 고정이므로, '
    '나머지 수요는 LNG가 맡을 수밖에 없다. 즉 비용이 아니라 무탄소 공급 설비의 양이 한계다.'),
   ('목표 도달에 필요한 대체량 (개략): ', f"부족분을 LNG 발전 배출계수({V(2035, 'MIN', 'LNG_tCO2_per_MWh'):.2f} t/MWh)로 환산하면 "
    f"2035년 약 {V(2035, 'MIN', 'need_TWh'):.0f} TWh, 2040년 약 {V(2040, 'MIN', 'need_TWh'):.0f} TWh의 LNG 발전을 "
    '무탄소 전력으로 바꿔야 한다. 이는 설비 확충(재생에너지·원자력·저장 등)이 있어야 가능하다.'),
   ('비용 민감도와의 관계: ', f"2035년은 비용 +10% 케이스({SENS[2035].CO2_Mt.iloc[-1]:.1f} Mt)가 이미 이 한계에 거의 닿아 있다. "
    f"2040년은 +10% 케이스({SENS[2040].CO2_Mt.iloc[-1]:.1f} Mt)에서 최소({fy(2040, 'CO2_Mt')} Mt)까지 약 "
    f"{SENS[2040].CO2_Mt.iloc[-1] - V(2040, 'MIN', 'CO2_Mt'):.0f} Mt의 추가 감축 여지가 있다.")])
T(summary_table(), font=7, caption='표 0-1. 결과 요약',
  fmts={'CO2 (Mt)': '{:,.2f}', '목표 대비 (Mt)': '{:+,.1f}', '시스템 비용 (조원, slack 제외)': '{:,.2f}', '비용 증가 (조원)': '{:,.2f}',
        '증가율 (%, slack 제외)': '{:,.1f}', '증가율 (%, 목적함수)': '{:,.1f}', '평균 감축비용 (천원/t)': '{:,.0f}', '연도': '{:.0f}'})
F('fig1_path_target.png', '그림 1. 비용 증가에 따른 최소 배출 경로와 감축목표 (점선: 비용 +10% → 배출 최소)', width=17)

doc.add_page_break()
H('1. 배출과 수급', 1)
F('fig2_co2_source.png', '그림 2. 배출원별 CO2 (기준 vs 배출 최소)', width=14)
F('fig3_mix.png', '그림 3. 발전원별 발전량 (빗금: 목표 도달을 위해 무탄소 전력으로 바꿔야 할 LNG 발전량, 개략)', width=14)
T(balance_table(), font=7, caption='표 1-1. 전력·열 수급과 배출 (TWh, Mt)', fmt='{:,.2f}')
B([('석탄 → LNG 전환: ', '배출 최소 운영에서는 석탄을 거의 끄고 LNG·CHP로 대체한다. 이 전환이 감축의 대부분이다.'),
   ('LNG 설비 여유: ', f"LNG 이용률은 2035년 {fy(2035, 'LNG_CF_%', '{:.0f}')}%, 2040년 {fy(2040, 'LNG_CF_%', '{:.0f}')}%로 설비는 남지만, "
    'LNG 자체가 배출원이라 더 쓴다고 배출이 줄지는 않는다.'),
   ('재생에너지: ', '출력제한이 일부 줄어 PV·풍력 발전량이 소폭 늘지만, 용량이 고정이라 증가 폭은 작다.'),
   ('보조공급: ', '열보조공급·전력 slack 사용량은 기준과 같다(실제 설비 운전만으로 줄인 결과).')])

doc.add_page_break()
H('2. 비용', 1)
T(cost_table(), font=7.5, caption='표 2-1. 비용 구성 (조원/년)', fmt='{:,.3f}')
B([('비용 증가의 대부분: ', '가스 연료비 증가가 석탄 연료비 감소보다 훨씬 크다(연료 전환 비용).'),
   ('평균 감축비용: ', f"2035년 {V(2035, 'MIN', 'avg_abate_kwon_t'):,.0f}천원/t, 2040년 {V(2040, 'MIN', 'avg_abate_kwon_t'):,.0f}천원/t "
    '(시스템 비용 증가 ÷ 배출 감소, 기준 대비).')])

H('3. 설비', 1)
P('설비 용량은 기준과 같은 입력이다. 확장 가능하도록 설정된 일부 보조 설비만 바뀌었다(표 3-1).')
T(capacity_table(), font=7.5, caption='표 3-1. 확장 가능 설비 용량 (GW / GWh)', fmt='{:,.2f}', fmts={'연도': '{:.0f}'})

doc.add_page_break()
H('부록. 방법·검증', 1)
B([('배출 최소화 방법: ', 'PyPSA 목적함수(기존 분석과 같은 비용 정의)에 탄소가격 λ×배출량을 더해 비용을 최소화했다. '
    'λ를 충분히 크게 두면 λ 이하 비용으로 가능한 감축이 모두 반영되므로, 배출량이 λ를 더 올려도 거의 줄지 않는 지점이 기술적 최소다. '
    '이렇게 얻은 해는 그 배출량을 달성하는 최소비용 운영이기도 하다. 모델 비용 단위가 천원이므로 λ 단위는 천원/tCO2.'),
   ('λ 상한의 이유: ', 'λ가 너무 크면 전력 slack(미공급)·열 보조공급(단가 약 500만원/MWh)을 늘려 배출을 "줄이는" 비현실적인 해가 나온다. '
    '그래서 보조공급 사용량이 기준과 같은 해만 채택했다(표 A-1, A-2).'),
   ('입력 조건: ', '두 연도 모두 비용 민감도 분석 때 저장한 기준 네트워크(B0.nc)에서 다시 최적화했다. '
    'λ=0으로 다시 푼 목적함수가 기준 결과(2035년 109.90830조원, 2040년 41.23387조원)와 일치한다. '
    '분석 도중 interface가 2040년 조건으로 바뀌어, interface로 만든 2035년 계산 두 건은 결과에서 제외했다.'),
   ('시스템 비용: ', 'PyPSA 목적함수에서 slack·Fallback 운전 페널티와 slack 설비 투자비를 후처리로 뺀 값.'),
   ('배출량: ', '발전기 연료 사용 × 배출계수(석탄 0.359, 가스 0.211 tCO2/MWh_fuel). CHP 연료는 CHP 연료공급 발전기로 포함된다.'),
   ('필요 대체량: ', '목표 부족분 ÷ 최소 배출 운영의 LNG 발전 배출계수. 실제 필요 설비량은 시간대별 수급·저장에 따라 달라지므로 개략치다.'),
   ('실행: ', r'run_co2_min.py (2035: results\대표경로\1105\5. 배출량최소화\2035, 2040: results\co2_min_2040_20261110_144137, '
    rf'검증: results\대표경로\1105\5. 배출량최소화\2040), 보고서 make_co2min_report.py → {OUT}')])
fm = {'CO2 (Mt)': '{:,.2f}', '시스템 비용 증가 (%, slack 제외)': '{:,.1f}', '보조공급 변화 (GWh)': '{:+,.1f}',
      '탄소가격 λ (천원/t)': '{:,.0f}'}
T(CHK[2035], font=7.5, caption='표 A-1. 2035년 탄소가격별 결과와 보조공급 변화', fmts=fm)
T(CHK[2040], font=7.5, caption='표 A-2. 2040년 탄소가격별 결과와 보조공급 변화', fmts=fm)

path = os.path.join(OUT, 'CO2_최소배출_2035_2040_분석_보고서.docx')
doc.save(path)
print(f'완료: {path}')
pd.set_option('display.width', 250); pd.set_option('display.max_columns', 30)
for k, v in TBL.items():
    print(k); print(v.round(2).to_string(index=False))
