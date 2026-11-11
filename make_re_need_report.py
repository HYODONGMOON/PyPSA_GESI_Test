"""
감축목표 달성에 필요한 추가 재생에너지 용량 보고서 (Word + 그림 + 표)

선행:
  python run_re_need.py --tag 2035 --b0-nc <2035 B0.nc> --target 88.3 --modes opt
  python run_re_need.py --tag 2040 --b0-nc <2040 B0.nc> --target 58.9 --modes opt
사용법:
  python make_re_need_report.py [--out reports/re_need_<날짜>]
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

from analyze_mga import analyze
from run_co2_cost_sweep import aux_capex, aux_gens
from run_mga_sweep import penalty_value

logging.disable(logging.WARNING)
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 9

ap = argparse.ArgumentParser()
ap.add_argument('--out', default=f"reports/re_need_{datetime.now():%Y%m%d}")
args = ap.parse_args()
OUT = args.out
os.makedirs(OUT, exist_ok=True)

YEARS = {
    2035: dict(target=88.3, B0='results/co2_price_search_2035_20261110_065714/B0.nc',
               MIN='results/_tmp_ref/co2min2035/CO2min-lam1000.nc', RE_dir='results/re_need_2035_20261110_162506'),
    2040: dict(target=58.9, B0='results/co2_cost_sweep_2040_20261109_145316/B0.nc',
               MIN='results/co2_min_2040_20261110_144137/CO2min-lam1000.nc', RE_dir='results/re_need_2040_20261110_163101'),
}
CASES = ('B0', 'MIN', 'RE')
CKR = {'B0': '기준 (비용최소)', 'MIN': '배출 최소 (설비 고정)', 'RE': '재생 증설 (목표 달성)'}
CSH = {'B0': '기준', 'MIN': '배출 최소', 'RE': '재생 증설'}
REG = ['GGD', 'ICN', 'GWD', 'CBD', 'CND', 'JBD', 'GBD', 'BSN', 'GND', 'JND', 'JJD']
RKR = {'BSN': '부산울산', 'CBD': '충북', 'CND': '대전세종충남', 'GBD': '대구경북', 'GGD': '서울경기', 'GND': '경남', 'GWD': '강원',
       'ICN': '인천', 'JBD': '전북', 'JJD': '제주', 'JND': '광주전남'}
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
    o['LNG_CF_%'] = 100 * o['LNG_power_TWh'] * 1e3 / (g.loc[lng, 'p_nom'].sum() / 1e3 * w.sum())
    o['aux_TWh'] = float(P[aux_gens(n)].mul(w, axis=0).sum().sum() / 1e6)
    o['power_slack_TWh'] = float(P[[x for x in g.index if 'slack' in x.lower() or 'failsafe' in x.lower()]]
                                 .mul(w, axis=0).sum().sum() / 1e6)
    o['Z'] = float(n.objective)
    o['penalty'] = penalty_value(n)
    o['slack_capex'] = aux_capex(n)
    o['C_sys'] = o['Z'] - o['penalty'] - o['slack_capex']
    return o


R, REGT = {}, {}
for y, cfg in YEARS.items():
    for c in CASES:
        print(f'[분석] {y} {c}', flush=True)
        n = pypsa.Network(os.path.join(cfg['RE_dir'], 'RE-opt.nc') if c == 'RE' else cfg[c])
        R[(y, c)] = metrics(n)
        del n
    b = R[(y, 'B0')]
    for c in CASES:
        r = R[(y, c)]
        r['dC_jo'] = (r['C_sys'] - b['C_sys']) / 1e9
        r['dC_%'] = 100 * (r['C_sys'] / b['C_sys'] - 1)
        r['dZ_%'] = 100 * (r['Z'] / b['Z'] - 1)
        r['dE_Mt'] = r['CO2_Mt'] - b['CO2_Mt']
        r['abate_kwon_t'] = 1e3 * r['dC_jo'] / -r['dE_Mt'] if r['dE_Mt'] < -1e-3 else np.nan
        r['RE_capex_jo'] = r.get('c_capex_PV', 0.0) + r.get('c_capex_WT', 0.0)
    rg = pd.read_csv(os.path.join(cfg['RE_dir'], 're_regional.csv'))
    rg = rg[rg.case == 'RE-opt']
    REGT[y] = {t: rg[rg.tech == t].set_index('region').reindex(REG) for t in ('PV', 'WT')}
    re = R[(y, 'RE')]
    re['PV_add_GW'] = float(REGT[y]['PV'].add_GW.sum())
    re['WT_add_GW'] = float(REGT[y]['WT'].add_GW.sum())
    re['PV_exist_GW'] = float(REGT[y]['PV'].exist_GW.sum())
    re['WT_exist_GW'] = float(REGT[y]['WT'].exist_GW.sum())
    re['RE_add_GW'] = re['PV_add_GW'] + re['WT_add_GW']
    re['dRE_TWh'] = re['PV_TWh'] + re['WT_TWh'] - b['PV_TWh'] - b['WT_TWh']
    re['RE_capex_per_GW'] = 1e4 * re['RE_capex_jo'] / re['RE_add_GW']


def V(y, c, k):
    return float(R[(y, c)].get(k, np.nan))


figs = {}


def save(fig, name, rect=(0, 0, 1, 1)):
    p = os.path.join(OUT, name)
    fig.tight_layout(rect=rect)
    fig.savefig(p, dpi=200)
    plt.close(fig)
    figs[name] = p


# 그림 1: 전국 PV·풍력 용량 (기존 + 추가)
fig, ax = plt.subplots(1, 2, figsize=(10.5, 3.9))
for a, y in zip(ax, YEARS):
    re = R[(y, 'RE')]
    xs = np.arange(2)
    ex = [re['PV_exist_GW'], re['WT_exist_GW']]
    ad = [re['PV_add_GW'], re['WT_add_GW']]
    a.bar(xs, ex, 0.55, color=['#C9A227', '#2C6E9B'], label='기존')
    a.bar(xs, ad, 0.55, bottom=ex, color=[C_PV, C_WT], hatch='//', edgecolor='white', label='추가')
    for i in range(2):
        a.text(i, ex[i] / 2, f'기존\n{ex[i]:.1f}', ha='center', va='center', color='white', fontsize=8.5)
        a.text(i, ex[i] + ad[i] + 4, f'+{ad[i]:.1f} GW\n(합계 {ex[i] + ad[i]:.1f})', ha='center', fontsize=8.5)
    a.set_xticks(xs, ['태양광', '풍력']); a.set_ylabel('GW')
    a.set_ylim(0, (re['PV_exist_GW'] + re['PV_add_GW']) * 1.2)
    a.set_title(f"({'a' if y == 2035 else 'b'}) {y}년: 추가 {re['RE_add_GW']:.1f} GW (목표 {YEARS[y]['target']} Mt)")
    a.grid(axis='y', alpha=0.3)
save(fig, 'fig1_re_add.png')

# 그림 2: 지역별 추가 용량 + 잠재량 사용률
fig, ax = plt.subplots(1, 2, figsize=(11.5, 4.3), sharey=True)
yy = np.arange(len(REG))[::-1]
for a, y in zip(ax, YEARS):
    pv, wt = REGT[y]['PV'], REGT[y]['WT']
    a.barh(yy + 0.2, pv.exist_GW, 0.4, color='#C9A227', label='PV 기존')
    a.barh(yy + 0.2, pv.add_GW, 0.4, left=pv.exist_GW, color=C_PV, hatch='//', edgecolor='white', label='PV 추가')
    a.barh(yy - 0.2, wt.exist_GW, 0.4, color='#2C6E9B', label='풍력 기존')
    a.barh(yy - 0.2, wt.add_GW, 0.4, left=wt.exist_GW, color=C_WT, hatch='//', edgecolor='white', label='풍력 추가')
    for i, r in enumerate(REG):
        for off, t in ((0.2, pv), (-0.2, wt)):
            if t.at[r, 'add_GW'] > 0.05:
                tot = t.at[r, 'opt_GW']
                a.text(tot + 0.5, yy[i] + off, f"+{t.at[r, 'add_GW']:.1f} ({100 * tot / t.at[r, 'potential_GW']:.0f}%)",
                       va='center', fontsize=7)
    a.set_title(f"({'a' if y == 2035 else 'b'}) {y}년 지역별 PV·풍력 (괄호: 잠재량 대비 총용량)")
    a.set_xlabel('GW'); a.grid(axis='x', alpha=0.3)
    a.set_xlim(0, max(pv.opt_GW.max(), wt.opt_GW.max()) * 1.35)
ax[0].set_yticks(yy, [RKR[r] for r in REG])
h, l_ = ax[0].get_legend_handles_labels()
fig.legend(h, l_, fontsize=8, loc='lower center', ncol=4, bbox_to_anchor=(0.5, -0.01))
save(fig, 'fig2_regional.png', rect=(0, 0.06, 1, 1))

# 그림 3: 배출원별 CO2 + 발전 믹스
fig, ax = plt.subplots(1, 2, figsize=(12, 4.2))
xs, labels = [], []
for i, (y, cfg) in enumerate(YEARS.items()):
    for j, c in enumerate(CASES):
        x = i * 3.6 + j
        bot = 0
        for key, nm, col in (('CO2_coal_Mt', '석탄', C_COAL), ('CO2_LNG_Mt', 'LNG 발전', C_LNG), ('CO2_CHPfuel_Mt', 'CHP 연료', C_CHP)):
            v = V(y, c, key)
            ax[0].bar(x, v, 0.75, bottom=bot, color=col, label=nm if (i, j) == (0, 0) else None)
            bot += v
        ax[0].text(x, bot + 3, f'{bot:.1f}', ha='center', fontsize=8)
        xs.append(x); labels.append(f'{y}\n{CSH[c]}')
    ax[0].hlines(cfg['target'], i * 3.6 - 0.5, i * 3.6 + 2.5, color='#2E7D32', ls='--', lw=1.3,
                 label='감축목표' if i == 0 else None)
ax[0].set_xticks(xs, labels, fontsize=8); ax[0].set_ylabel('Mt'); ax[0].legend(fontsize=7.5, ncol=2, loc='upper right')
ax[0].set_ylim(0, 270); ax[0].grid(axis='y', alpha=0.3); ax[0].set_title('(a) 배출원별 CO2')
mix = [('석탄', 'coal_TWh', C_COAL), ('LNG 발전', 'LNG_power_TWh', C_LNG), ('CHP 발전', 'CHP_power_TWh', C_CHP),
       ('원자력', 'nuclear_TWh', C_NU), ('PV', 'PV_TWh', C_PV), ('풍력', 'WT_TWh', C_WT)]
for i, y in enumerate(YEARS):
    for j, c in enumerate(CASES):
        x = i * 3.6 + j
        bot = 0
        for nm, k, col in mix:
            v = V(y, c, k)
            ax[1].bar(x, v, 0.75, bottom=bot, color=col, label=nm if (i, j) == (0, 0) else None)
            bot += v
        cur = V(y, c, 'curtail_TWh')
        ax[1].bar(x, cur, 0.75, bottom=bot, color='white', edgecolor='#999', hatch='..',
                  label='출력제한' if (i, j) == (0, 0) else None)
ax[1].set_xticks(xs, labels, fontsize=8); ax[1].set_ylabel('TWh')
ax[1].legend(fontsize=7.5, loc='center left', bbox_to_anchor=(1.01, 0.5)); ax[1].grid(axis='y', alpha=0.3)
ax[1].set_title('(b) 발전원별 발전량 (점무늬: 출력제한)')
save(fig, 'fig3_co2_mix.png')

# 그림 4: 비용 구성
CG = [('석탄 연료비', ['c_opex_coal'], C_COAL), ('가스 연료비', ['c_opex_gas'], C_LNG), ('CHP 운전비', ['c_opex_CHP'], C_CHP),
      ('원자력 연료비', ['c_opex_nuclear'], C_NU), ('PV 투자비 (추가분)', ['c_capex_PV'], C_PV), ('풍력 투자비 (추가분)', ['c_capex_WT'], C_WT)]


def cg(y, c, ks):
    return sum(float(R[(y, c)].get(k, 0.0) or 0.0) for k in ks)


fig, ax = plt.subplots(figsize=(9.5, 4))
for i, y in enumerate(YEARS):
    for j, c in enumerate(CASES):
        x = i * 3.6 + j
        bot = 0
        for nm, ks, col in CG:
            v = cg(y, c, ks)
            ax.bar(x, v, 0.75, bottom=bot, color=col, label=nm if (i, j) == (0, 0) else None)
            bot += v
        oth = V(y, c, 'C_sys') / 1e9 - bot
        ax.bar(x, oth, 0.75, bottom=bot, color='#BDBDBD', label='기타' if (i, j) == (0, 0) else None)
        ax.text(x, V(y, c, 'C_sys') / 1e9 + 0.6, f"{V(y, c, 'C_sys') / 1e9:.1f}", ha='center', fontsize=8)
ax.set_xticks(xs, labels, fontsize=8); ax.set_ylabel('조원/년')
ax.legend(fontsize=7.5, loc='center left', bbox_to_anchor=(1.01, 0.5)); ax.grid(axis='y', alpha=0.3)
ax.set_title('시스템 비용 구성 (slack 비용 제외)')
save(fig, 'fig4_cost.png')


# ───────────── 표 ─────────────
def summary_table():
    rows = []
    for y, cfg in YEARS.items():
        for c in CASES:
            rows.append({'연도': y, '경우': CKR[c], 'CO2 (Mt)': V(y, c, 'CO2_Mt'), '목표 대비 (Mt)': V(y, c, 'CO2_Mt') - cfg['target'],
                         'PV 추가 (GW)': V(y, c, 'PV_add_GW') if c == 'RE' else 0.0,
                         '풍력 추가 (GW)': V(y, c, 'WT_add_GW') if c == 'RE' else 0.0,
                         '시스템 비용 (조원)': V(y, c, 'C_sys') / 1e9, '비용 증가 (조원)': V(y, c, 'dC_jo'),
                         '증가율 (%, slack 제외)': V(y, c, 'dC_%'), '평균 감축비용 (천원/t)': V(y, c, 'abate_kwon_t')})
    return pd.DataFrame(rows)


def regional_table(y):
    pv, wt = REGT[y]['PV'], REGT[y]['WT']
    df = pd.DataFrame({'지역': [RKR[r] for r in REG],
                       'PV 기존': pv.exist_GW.values, 'PV 추가': pv.add_GW.values, 'PV 합계': pv.opt_GW.values,
                       'PV 잠재량': pv.potential_GW.values, 'PV 잠재량 사용 (%)': 100 * (pv.opt_GW / pv.potential_GW).values,
                       '풍력 기존': wt.exist_GW.values, '풍력 추가': wt.add_GW.values, '풍력 합계': wt.opt_GW.values,
                       '풍력 잠재량': wt.potential_GW.values, '풍력 잠재량 사용 (%)': 100 * (wt.opt_GW / wt.potential_GW).values})
    tot = {'지역': '합계'}
    for c in df.columns[1:]:
        tot[c] = df[c].sum()
    tot['PV 잠재량 사용 (%)'] = 100 * tot['PV 합계'] / tot['PV 잠재량']
    tot['풍력 잠재량 사용 (%)'] = 100 * tot['풍력 합계'] / tot['풍력 잠재량']
    return pd.concat([df, pd.DataFrame([tot])], ignore_index=True)


def balance_table():
    its = [('석탄 발전', 'coal_TWh'), ('LNG 발전', 'LNG_power_TWh'), ('CHP 발전', 'CHP_power_TWh'), ('원자력', 'nuclear_TWh'),
           ('PV 발전', 'PV_TWh'), ('풍력 발전', 'WT_TWh'), ('출력제한', 'curtail_TWh'), ('출력제한률 (%)', 'curtail_%'),
           ('PV 이용률 (%)', 'PV_CF_%'), ('풍력 이용률 (%)', 'WT_CF_%'), ('LNG 이용률 (%)', 'LNG_CF_%'),
           ('히트펌프 열', 'HP_heat_TWh'), ('CHP 열', 'CHP_heat_TWh'), ('수전해 전력', 'electrolyser_elec_TWh'),
           ('ESS 충전', 'ESS_charge_TWh'), ('HVDC 수송량', 'HVDC_TWh'), ('열보조공급', 'heat_fallback_TWh'),
           ('전력 slack', 'power_slack_TWh'), ('ESS 출력 (GW)', 'ESS_GW'), ('수전해 (GW)', 'electrolyser_GW'),
           ('CO2 석탄 (Mt)', 'CO2_coal_Mt'), ('CO2 LNG 발전 (Mt)', 'CO2_LNG_Mt'), ('CO2 CHP 연료 (Mt)', 'CO2_CHPfuel_Mt'),
           ('CO2 합계 (Mt)', 'CO2_Mt')]
    df = pd.DataFrame({'항목 (TWh)': [a for a, _ in its]})
    for y in YEARS:
        for c in CASES:
            df[f'{y} {CSH[c]}'] = [V(y, c, k) for _, k in its]
    return df


def cost_table():
    rows = [nm for nm, _, _ in CG] + ['기타', '시스템 비용 합계', '(제외) slack·Fallback 운전 페널티', '(제외) slack 설비 투자비']
    df = pd.DataFrame({'비용 항목 (조원/년)': rows})
    for y in YEARS:
        for c in CASES:
            vals = [cg(y, c, ks) for _, ks, _ in CG]
            tot = V(y, c, 'C_sys') / 1e9
            df[f'{y} {CSH[c]}'] = vals + [tot - sum(vals), tot, V(y, c, 'penalty') / 1e9, V(y, c, 'slack_capex') / 1e9]
    return df


TBL = {'표1_요약': summary_table(), '표2_지역별_2035': regional_table(2035), '표3_지역별_2040': regional_table(2040),
       '표4_수급': balance_table(), '표5_비용': cost_table()}
with pd.ExcelWriter(os.path.join(OUT, '재생에너지_추가용량_2035_2040_표.xlsx')) as xw:
    for k, v in TBL.items():
        v.to_excel(xw, sheet_name=k, index=False)
    pd.DataFrame({f'{y} {c}': R[(y, c)] for y in YEARS for c in CASES}).drop(index=['case'], errors='ignore') \
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
            if isinstance(v, (float, np.floating, int, np.integer)) and not isinstance(v, bool):
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


def top_regions(y, k=4):
    items = [(REGT[y][t].at[r, 'add_GW'], r, t) for r in REG for t in ('PV', 'WT') if REGT[y][t].at[r, 'add_GW'] > 0.05]
    items.sort(reverse=True)
    return ', '.join(f"{RKR[r]} {'PV' if t == 'PV' else '풍력'} +{v:.1f}" for v, r, t in items[:k])


def full_pot(y):
    out = []
    for r in REG:
        for t in ('PV', 'WT'):
            d = REGT[y][t]
            if d.at[r, 'add_GW'] > 0.05 and d.at[r, 'opt_GW'] >= 0.98 * d.at[r, 'potential_GW']:
                out.append(f"{RKR[r]} {'PV' if t == 'PV' else '풍력'}")
    return ', '.join(out)


doc.add_heading('2035·2040년 감축목표 달성에 필요한 추가 재생에너지 용량', 0)
P('설비 고정 조건의 최소 배출로도 남는 부족분을 PV·풍력 증설로 메울 때의 최소비용 용량 조합', size=12)
P(f'작성일: {datetime.now():%Y-%m-%d}   |   모형: PyPSA 0.31.1 + linopy + CPLEX Barrier, 8,784시간 단일연도', size=9)

H = doc.add_heading
H('개요', 1)
P('앞선 배출량 최소화 분석에서 현재 설비로는 감축목표(2035년 88.3 Mt, 2040년 58.9 Mt)에 도달할 수 없음을 확인했다. '
  '이번 분석은 각 연도의 기존 분석 조건에서 PV·풍력만 추가 증설을 허용하고, 배출량이 목표 이하가 되도록 하는 최소비용 '
  '재생에너지 용량을 구한다. 2035년과 2040년은 서로 연계하지 않고 각각 독립적으로 계산했다.')

H('핵심 결과', 2)
r35, r40 = R[(2035, 'RE')], R[(2040, 'RE')]
B([('필요 추가 용량: ', f"2035년 {r35['RE_add_GW']:.1f} GW(PV {r35['PV_add_GW']:.1f}, 풍력 {r35['WT_add_GW']:.1f}), "
    f"2040년 {r40['RE_add_GW']:.1f} GW(PV {r40['PV_add_GW']:.1f}, 풍력 {r40['WT_add_GW']:.1f}). "
    '최소비용 조합은 두 연도 모두 PV 중심(약 90%)이다.'),
   ('증설 후 총용량: ', f"2035년 PV {r35['PV_exist_GW'] + r35['PV_add_GW']:.1f}·풍력 {r35['WT_exist_GW'] + r35['WT_add_GW']:.1f} GW, "
    f"2040년 PV {r40['PV_exist_GW'] + r40['PV_add_GW']:.1f}·풍력 {r40['WT_exist_GW'] + r40['WT_add_GW']:.1f} GW."),
   ('주요 증설 지역: ', f"2035년 {top_regions(2035)}, 2040년 {top_regions(2040)} GW. "
    + (f"잠재량을 모두 채우는 지역: 2035년 {full_pot(2035) or '없음'}, 2040년 {full_pot(2040) or '없음'}." )),
   ('필요 비용: ', f"시스템 비용(slack 제외)은 2035년 {V(2035, 'RE', 'C_sys') / 1e9:.2f}조원(+{V(2035, 'RE', 'dC_jo'):.2f}조원, "
    f"+{V(2035, 'RE', 'dC_%'):.1f}%), 2040년 {V(2040, 'RE', 'C_sys') / 1e9:.2f}조원(+{V(2040, 'RE', 'dC_jo'):.2f}조원, "
    f"+{V(2040, 'RE', 'dC_%'):.1f}%). 이 중 추가 PV·풍력의 연간 투자비가 2035년 {r35['RE_capex_jo']:.2f}조원, "
    f"2040년 {r40['RE_capex_jo']:.2f}조원이다."),
   ('설비 고정 최소 배출과 비교: ', f"2035년은 설비 고정 최소 배출({V(2035, 'MIN', 'CO2_Mt'):.1f} Mt, {V(2035, 'MIN', 'C_sys') / 1e9:.2f}조원)보다 "
    f"배출은 {V(2035, 'MIN', 'CO2_Mt') - V(2035, 'RE', 'CO2_Mt'):.1f} Mt 적고 비용은 {V(2035, 'MIN', 'C_sys') / 1e9 - V(2035, 'RE', 'C_sys') / 1e9:.2f}조원 낮다. "
    f"2040년은 배출 {V(2040, 'MIN', 'CO2_Mt') - V(2040, 'RE', 'CO2_Mt'):.1f} Mt를 더 줄이는 데 "
    f"{V(2040, 'RE', 'C_sys') / 1e9 - V(2040, 'MIN', 'C_sys') / 1e9:.2f}조원이 더 든다. 가스 연료비 대신 재생 투자비를 쓰는 구조다."),
   ('부작용: ', f"출력제한이 2035년 {V(2035, 'B0', 'curtail_TWh'):.1f} → {V(2035, 'RE', 'curtail_TWh'):.1f} TWh"
    f"({V(2035, 'RE', 'curtail_%'):.1f}%), 2040년 {V(2040, 'B0', 'curtail_TWh'):.1f} → {V(2040, 'RE', 'curtail_TWh'):.1f} TWh"
    f"({V(2040, 'RE', 'curtail_%'):.1f}%)로 늘어난다. 저장 설비는 기존 조건 그대로이므로, 저장을 함께 늘리면 필요한 재생 용량이 줄 수 있다.")])
T(summary_table(), font=7, caption='표 0-1. 결과 요약',
  fmts={'연도': '{:.0f}', 'CO2 (Mt)': '{:,.2f}', '목표 대비 (Mt)': '{:+,.1f}', '시스템 비용 (조원)': '{:,.2f}', '비용 증가 (조원)': '{:,.2f}',
        '증가율 (%, slack 제외)': '{:,.1f}', '평균 감축비용 (천원/t)': '{:,.0f}'})
F('fig1_re_add.png', '그림 1. 감축목표 달성에 필요한 PV·풍력 용량 (빗금: 추가분)', width=15)

doc.add_page_break()
H('1. 지역별 증설', 1)
F('fig2_regional.png', '그림 2. 지역별 PV·풍력 용량 (빗금: 추가분, 괄호: 잠재량 대비 총용량)', width=17)
P('증설은 수요가 큰 수도권(서울경기·인천)과 충청권, 그리고 PV 비용이 낮은 지역에 집중된다. 지역별 투자비 계수와 송전 제약을 반영한 '
  '최소비용 배치이며, 잠재량은 기존 분석에서 쓴 지역별 기술적 잠재량이다.')
fr = {c: '{:,.1f}' for c in TBL['표2_지역별_2035'].columns if c != '지역'}
fr.update({'PV 잠재량 사용 (%)': '{:,.0f}', '풍력 잠재량 사용 (%)': '{:,.0f}'})
T(TBL['표2_지역별_2035'], font=6.5, caption='표 1-1. 2035년 지역별 PV·풍력 (GW)', fmts=fr)
T(TBL['표3_지역별_2040'], font=6.5, caption='표 1-2. 2040년 지역별 PV·풍력 (GW)', fmts=fr)

doc.add_page_break()
H('2. 배출과 수급', 1)
F('fig3_co2_mix.png', '그림 3. 배출원별 CO2(a)와 발전원별 발전량(b): 기준 / 설비 고정 배출 최소 / 재생 증설', width=17.5)
B([('석탄: ', '재생 증설 경우에도 석탄은 거의 쓰지 않는다. 늘어난 재생 발전이 LNG 발전의 일부를 대신해 배출이 목표까지 내려간다.'),
   ('LNG: ', f"재생 증설 경우 LNG 발전은 2035년 {V(2035, 'RE', 'LNG_power_TWh'):.0f} TWh, 2040년 {V(2040, 'RE', 'LNG_power_TWh'):.0f} TWh로 "
    f"설비 고정 배출 최소({V(2035, 'MIN', 'LNG_power_TWh'):.0f}, {V(2040, 'MIN', 'LNG_power_TWh'):.0f} TWh)보다 적다. "
    '재생 발전이 적은 시간대의 수요는 여전히 LNG가 맡는다.'),
   ('재생 발전량: ', f"PV·풍력 발전은 기준 대비 2035년 +{r35['dRE_TWh']:.0f} TWh, 2040년 +{r40['dRE_TWh']:.0f} TWh 늘었다. "
    f"추가 용량에 비해 발전량 증가가 작은 것은 출력제한이 늘고 PV 이용률이 떨어지기 때문이다."),
   ('보조공급: ', '열보조공급·전력 slack은 늘지 않았다(전력 slack은 오히려 소폭 감소). 미공급으로 배출을 줄인 결과가 아니다.')])
T(balance_table(), font=6.5, caption='표 2-1. 수급과 배출 (TWh, Mt)', fmt='{:,.1f}')

doc.add_page_break()
H('3. 비용', 1)
F('fig4_cost.png', '그림 4. 시스템 비용 구성 (slack 비용 제외)', width=15)
T(cost_table(), font=7, caption='표 3-1. 비용 구성 (조원/년)', fmt='{:,.2f}')
B([('비용 구조: ', '재생 증설 경우 비용 증가분은 PV·풍력 연간 투자비와 가스 연료비 증가로 이루어진다. 석탄 연료비 감소가 일부를 상쇄한다.'),
   ('GW당 연간 투자비: ', f"추가 재생 1 GW당 2035년 {r35['RE_capex_per_GW']:,.0f}억원/년, 2040년 {r40['RE_capex_per_GW']:,.0f}억원/년 "
    '(입력 투자비 계수 기준, 지역·기술 가중 평균).'),
   ('평균 감축비용: ', f"기준 대비 2035년 {V(2035, 'RE', 'abate_kwon_t'):,.0f}천원/t, 2040년 {V(2040, 'RE', 'abate_kwon_t'):,.0f}천원/t.")])

H('부록. 방법·가정', 1)
B([('계산 방법: ', '각 연도 기준 네트워크(B0.nc, 기존 분석 조건)에서 PV·풍력의 증설 상한만 지역별 기술적 잠재량으로 풀고 '
    '(하한은 기존 용량), 배출량 ≤ 목표 제약 아래 기존 목적함수(비용)를 최소화했다. 그 외 설비 조건은 기존과 같다.'),
   ('배출량: ', '발전기 연료 사용 × 배출계수(석탄 0.359, 가스 0.211 tCO2/MWh_fuel), CHP 연료 포함. 앞선 분석과 같은 산식.'),
   ('보조공급 제한: ', '열·전력 보조공급(Fallback·slack)의 연간 사용량은 발전기별로 기준 + 1 GWh 이하로 제한해, 미공급으로 배출을 줄이지 못하게 했다.'),
   ('시스템 비용: ', 'PyPSA 목적함수에서 slack·Fallback 운전 페널티와 slack 설비 투자비를 후처리로 뺀 값. 추가 PV·풍력의 연간화 투자비 포함.'),
   ('감축목표: ', '2035년 88.3 Mt, 2040년 58.9 Mt (앞선 분석과 같은 값).'),
   ('해석상 유의: ', '결과는 저장·송전·수요 조건이 기존과 같을 때의 최소비용 재생 용량이다. 저장 확대, 송전 보강, 원자력 등 다른 무탄소 전원을 함께 '
    '고려하면 필요한 재생 용량과 비용은 달라진다. 지역 배치는 모형의 투자비 계수와 잠재량에 따른 결과로, 실제 입지 여건은 별도 검토가 필요하다.'),
   ('실행: ', f"run_re_need.py → {YEARS[2035]['RE_dir']}, {YEARS[2040]['RE_dir']}, 보고서 make_re_need_report.py → {OUT}")])

path = os.path.join(OUT, '재생에너지_추가용량_2035_2040_분석_보고서.docx')
doc.save(path)
print(f'완료: {path}')
pd.set_option('display.width', 250); pd.set_option('display.max_columns', 30)
for k, v in TBL.items():
    print(k); print(v.round(2).to_string(index=False))
