"""
CO2 제약 해제·설비 고정 조건의 비용 상한별 최소 배출 대안 분석 보고서 (Word + 그림 + 표)

선행:
  python run_co2_price_search.py --reuse-b0 <B0.nc 폴더>
사용법:
  python make_co2cost_report.py --run-dir results/co2_price_search_2040_... [--out reports/co2cost2040_<날짜>] [--year 2040]
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

logging.disable(logging.WARNING)
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False
plt.rcParams['font.size'] = 9

ap = argparse.ArgumentParser()
ap.add_argument('--run-dir', required=True)
ap.add_argument('--out')
ap.add_argument('--year', type=int, default=2040)
ap.add_argument('--ref-dir')
args = ap.parse_args()
args.out = args.out or f"reports/co2cost{args.year}_{datetime.now():%Y%m%d}"
args.ref_dir = args.ref_dir or rf'results\대표경로\1105\3. 설비고정+CO2CAPX\{args.year}'
OUT = args.out
os.makedirs(OUT, exist_ok=True)

REGION_KR = {'BSN': '부산울산', 'CBD': '충북', 'CND': '대전세종충남', 'GBD': '대구경북', 'GGD': '서울경기',
             'GND': '경남', 'GWD': '강원', 'ICN': '인천', 'JBD': '전북', 'JJD': '제주', 'JND': '광주전남'}
REG = ['GGD', 'ICN', 'GWD', 'CBD', 'CND', 'JBD', 'GBD', 'BSN', 'GND', 'JND', 'JJD']
summ = pd.read_csv(os.path.join(args.run_dir, 'summary.csv'))
summ = summ[summ.ok.astype(bool)].drop_duplicates('case', keep='last').set_index('case')
CASES = [c for c in ['B0', 'CO2min-0', 'CO2min-1', 'CO2min-3', 'CO2min-5', 'CO2min-10'] if c in summ.index]
CASE_KR = {'B0': '기준 (비용최소)', 'CO2min-0': '+0%', 'CO2min-1': '+1%', 'CO2min-3': '+3%',
           'CO2min-5': '+5%', 'CO2min-10': '+10%'}
EPS = {c: float(summ.at[c, 'eps_pct']) for c in CASES}
ALT = [c for c in CASES if c != 'B0']
C_COAL, C_LNG, C_CHP, C_NU, C_PV, C_WT, C_FC, C_OTH = '#5D4037', '#E07B39', '#8E5EA2', '#6C757D', '#F2B134', '#3C8DBC', '#2E7D32', '#BDBDBD'
CCOL = {'B0': '#9E9E9E', 'CO2min-0': '#BDD7EE', 'CO2min-1': '#8DB4E2', 'CO2min-3': '#2E75B6', 'CO2min-5': '#1F4E79',
        'CO2min-10': '#0B2545'}


def extra(n):
    """배출원별 CO2·발전원별 지역 발전량·확장설비 변화"""
    w = n.snapshot_weightings.generators
    g = n.generators
    em = n.carriers['co2_emissions']
    eff = n.get_switchable_as_dense('Generator', 'efficiency')
    P = n.generators_t.p.mul(w, axis=0).sum()
    out = {}
    coal = g.index[g.carrier == 'coal']
    lng = [x for x in g.index[g.carrier == 'gas'] if x.endswith('_LNG')]
    fuel = [x for x in g.index[g.carrier == 'gas'] if 'Fuel_Supply' in x]
    for key, idx, car in (('coal', coal, 'coal'), ('LNG', lng, 'gas'), ('CHPfuel', fuel, 'gas')):
        out[f'CO2_{key}_Mt'] = float((n.generators_t.p[idx] / eff[idx]).mul(w, axis=0).sum().sum() * em[car] / 1e6)
    out['coal_TWh'] = float(P[coal].sum() / 1e6)
    out['coal_GW'] = g.loc[coal, 'p_nom'].sum() / 1e3
    out['coal_CF_%'] = 100 * out['coal_TWh'] * 1e3 / (out['coal_GW'] * w.sum())
    out['LNG_GW'] = g.loc[lng, 'p_nom'].sum() / 1e3
    sl = [x for x in g.index if 'slack' in x.lower() or 'failsafe' in x.lower()]
    out['power_slack_TWh'] = float(P[sl].sum() / 1e6)
    out['objective_full'] = float(n.objective)
    gc = n.global_constraints
    out['CO2_cap_Mt'] = float(gc.at['CO2_limit', 'constant']) / 1e6 if 'CO2_limit' in gc.index else np.nan
    reg = {}
    for r in REG:
        reg[r] = {'coal_TWh': float(P[[x for x in coal if x.startswith(r + '_')]].sum() / 1e6),
                  'LNG_TWh': float(P[[x for x in lng if x.startswith(r + '_')]].sum() / 1e6)}
    caps = {}
    for comp, attr in (('Generator', 'p_nom'), ('Link', 'p_nom'), ('Store', 'e_nom')):
        df = n.static(comp)
        ext = df.index[df[f'{attr}_extendable'].astype(bool)]
        for nm in ext:
            k = {'solar': 'PV', 'wind': 'WT'}.get(df.at[nm, 'carrier'], kind(nm)) if comp == 'Generator' else kind(nm)
            u = 'GWh' if comp == 'Store' else 'GW'
            caps[(k, u)] = caps.get((k, u), 0.0) + float(df.at[nm, f'{attr}_opt']) / 1e3
            caps[(k + '|exist', u)] = caps.get((k + '|exist', u), 0.0) + float(df.at[nm, attr]) / 1e3
    return out, reg, caps


nat, REGD, CAPS = {}, {}, {}
for c in CASES:
    print(f'[분석] {c}', flush=True)
    n = pypsa.Network(os.path.join(args.run_dir, f'{c}.nc'))
    o, _ = analyze(n, c)
    e, rg, cp = extra(n)
    o.update(e)
    o['cost_sys'] = o['cost_total_check'] - o.get('c_opex_penalty', 0.0) - o.get('c_capex_slack', 0.0)
    nat[c], REGD[c], CAPS[c] = o, rg, cp
    del n
nat = pd.DataFrame(nat).T
nat = nat.join(summ[[k for k in ('eps_pct', 'cost_cap', 'C_sys', 'cost_increase_pct', 'objective_increase_pct',
                                 'lambda_won_per_t', 'stage1_CO2_Mt', 'slack_capex', 'penalty_cost') if k in summ]])
nat = nat.rename(columns={'lambda_won_per_t': 'lambda_kwon_t'})


def V(c, k):
    return float(nat.at[c, k])


def d(c, k, ref='B0'):
    return V(c, k) - V(ref, k)


C0, E0 = V('B0', 'C_sys') / 1e9, V('B0', 'CO2_Mt')
for c in CASES:
    nat.at[c, 'C_sys_jo'] = V(c, 'C_sys') / 1e9
    nat.at[c, 'dC_jo'] = V(c, 'C_sys') / 1e9 - C0
    nat.at[c, 'dE_Mt'] = V(c, 'CO2_Mt') - E0
    nat.at[c, 'dE_%'] = 100 * (V(c, 'CO2_Mt') / E0 - 1)
    nat.at[c, 'avg_abate_kwon_t'] = (1e3 * nat.at[c, 'dC_jo'] / -nat.at[c, 'dE_Mt']) if nat.at[c, 'dE_Mt'] < -1e-3 else np.nan
for i, c in enumerate(CASES[1:], start=1):
    p_ = CASES[i - 1]
    dE = V(c, 'CO2_Mt') - V(p_, 'CO2_Mt')
    dC = V(c, 'C_sys_jo') - V(p_, 'C_sys_jo')
    nat.at[c, 'step_abate_kwon_t'] = 1e3 * dC / -dE if dE < -1e-3 else np.nan

figs = {}


def save(fig, name, rect=(0, 0, 1, 1)):
    p = os.path.join(OUT, name)
    fig.tight_layout(rect=rect)
    fig.savefig(p, dpi=200)
    plt.close(fig)
    figs[name] = p


# 그림 1: 비용 증가 vs 배출
fig, ax = plt.subplots(1, 2, figsize=(11, 3.9))
x = [V(c, 'cost_increase_pct') for c in CASES]; y = [V(c, 'CO2_Mt') for c in CASES]
ax[0].plot(x, y, '-o', color='#1F4E79')
for c, xi, yi in zip(CASES, x, y):
    ax[0].annotate(f"{CASE_KR[c]}\n{yi:.1f} Mt", (xi, yi), textcoords='offset points',
                   xytext=(-4, -32) if c == 'CO2min-0' else (6, 4), fontsize=7.5)
ax[0].set_xlabel('기준 대비 시스템 비용 증가율 (%, slack 비용 제외)'); ax[0].set_ylabel('CO2 배출량 (Mt)')
ax[0].grid(alpha=0.3); ax[0].set_title('(a) 비용 증가 허용률별 최소 배출량')
xx = np.arange(len(CASES)); bot = np.zeros(len(CASES))
for key, nm, col in (('CO2_coal_Mt', '석탄', C_COAL), ('CO2_LNG_Mt', 'LNG 발전', C_LNG), ('CO2_CHPfuel_Mt', 'CHP 연료', C_CHP)):
    v = np.array([V(c, key) for c in CASES])
    ax[1].bar(xx, v, 0.6, bottom=bot, color=col, label=nm); bot += v
for i, c in enumerate(CASES):
    ax[1].text(i, bot[i] + 2, f'{bot[i]:.1f}', ha='center', fontsize=8)
ax[1].set_xticks(xx, [CASE_KR[c] for c in CASES], fontsize=8); ax[1].set_ylabel('Mt')
ax[1].legend(fontsize=8, loc='upper right'); ax[1].grid(axis='y', alpha=0.3); ax[1].set_title('(b) 배출원별 CO2')
ax[1].set_ylim(0, max(bot) * 1.15)
save(fig, 'fig1_pareto.png')

# 그림 2: 발전 믹스
mix = [('석탄', 'coal_TWh', C_COAL), ('LNG 발전', 'LNG_power_TWh', C_LNG), ('CHP 발전', 'CHP_power_TWh', C_CHP),
       ('원자력', 'nuclear_TWh', C_NU), ('PV', 'PV_TWh', C_PV), ('풍력', 'WT_TWh', C_WT),
       ('연료전지', 'fuelcell_power_TWh', C_FC)]
fig, ax = plt.subplots(1, 2, figsize=(11.5, 3.9))
bot = np.zeros(len(CASES))
for nm, k, col in mix:
    v = np.array([V(c, k) for c in CASES])
    ax[0].bar(xx, v, 0.6, bottom=bot, color=col, label=nm); bot += v
ax[0].set_xticks(xx, [CASE_KR[c] for c in CASES], fontsize=8); ax[0].set_ylabel('TWh')
ax[0].legend(fontsize=7.5, loc='lower right', ncol=2); ax[0].grid(axis='y', alpha=0.3); ax[0].set_title('(a) 발전원별 발전량')
items = [('석탄', 'coal_TWh'), ('LNG 발전', 'LNG_power_TWh'), ('CHP 발전', 'CHP_power_TWh'), ('원자력', 'nuclear_TWh'),
         ('재생 발전', 'RE_TWh'), ('출력제한', 'curtail_TWh'), ('연료전지', 'fuelcell_power_TWh'),
         ('수전해 전력', 'electrolyser_elec_TWh'), ('히트펌프 전력', 'HP_elec_TWh'), ('ESS 충전', 'ESS_charge_TWh')]
nat['RE_TWh'] = nat['PV_TWh'] + nat['WT_TWh']
xi = np.arange(len(items)); wd = 0.8 / len(ALT)
for j, c in enumerate(ALT):
    ax[1].bar(xi - 0.4 + wd * (j + 0.5), [d(c, k) for _, k in items], wd, color=CCOL[c], label=CASE_KR[c])
ax[1].axhline(0, color='k', lw=0.6)
ax[1].set_xticks(xi, [a for a, _ in items], rotation=30, ha='right', fontsize=8)
ax[1].set_ylabel('기준 대비 변화 (TWh)'); ax[1].legend(fontsize=7.5, ncol=3); ax[1].grid(axis='y', alpha=0.3)
ax[1].set_title('(b) 기준 대비 연간 수급 변화')
save(fig, 'fig2_mix.png')

# 그림 3: 지역별 석탄·LNG 발전
fig, ax = plt.subplots(1, 2, figsize=(11.5, 4), sharey=True)
yy = np.arange(len(REG))[::-1]
show = ['B0'] + [c for c in ('CO2min-1', 'CO2min-5', 'CO2min-10') if c in CASES]
hh = 0.8 / len(show)
for k, (key, ttl) in enumerate((('coal_TWh', '(a) 석탄 발전 (TWh)'), ('LNG_TWh', '(b) LNG 발전 (TWh)'))):
    for j, c in enumerate(show):
        ax[k].barh(yy + 0.4 - hh * (j + 0.5), [REGD[c][r][key] for r in REG], hh, color=CCOL[c], label=CASE_KR[c])
    ax[k].set_title(ttl); ax[k].grid(axis='x', alpha=0.3)
ax[0].set_yticks(yy, [REGION_KR[r] for r in REG]); ax[1].legend(fontsize=8)
save(fig, 'fig3_regional.png')

# 그림 4: 감축비용
fig, ax = plt.subplots(figsize=(8.5, 3.6))
xs = np.arange(len(ALT))
ax.bar(xs - 0.2, [nat.at[c, 'avg_abate_kwon_t'] for c in ALT], 0.4, color='#2E75B6', label='평균 (기준 대비)')
ax.bar(xs + 0.2, [nat.at[c, 'step_abate_kwon_t'] for c in ALT], 0.4, color='#C0392B', label='구간 한계 (직전 단계 대비)')
for i, c in enumerate(ALT):
    for off, k in ((-0.2, 'avg_abate_kwon_t'), (0.2, 'step_abate_kwon_t')):
        v = nat.at[c, k]
        if pd.notna(v):
            ax.text(i + off, v, f'{v:,.0f}', ha='center', va='bottom', fontsize=7.5)
ax.set_xticks(xs, [CASE_KR[c] for c in ALT]); ax.set_ylabel('천원/tCO2'); ax.legend(fontsize=8); ax.grid(axis='y', alpha=0.3)
ax.set_title('CO2 감축비용 (시스템 비용 증가 ÷ 배출 감소)')
save(fig, 'fig4_abatement.png')


# ───────────── 표 ─────────────
def summary_table():
    rows = []
    for c in CASES:
        rows.append({'시나리오': CASE_KR[c], '비용 상한 (목적함수 기준)': '—' if c == 'B0' else f"+{EPS[c]:g}%",
                     '목적함수 실제 증가율 (%)': V(c, 'objective_increase_pct'),
                     '시스템 비용 (조원, slack 제외)': V(c, 'C_sys_jo'), '비용 증가 (조원)': V(c, 'dC_jo'),
                     'slack 제외 증가율 (%)': V(c, 'cost_increase_pct'), 'CO2 (Mt)': V(c, 'CO2_Mt'),
                     '기준 대비 CO2 (Mt)': V(c, 'dE_Mt'), '기준 대비 CO2 (%)': V(c, 'dE_%'),
                     '석탄 발전 (TWh)': V(c, 'coal_TWh'), 'LNG 발전 (TWh)': V(c, 'LNG_power_TWh'),
                     '평균 감축비용 (천원/t)': nat.at[c, 'avg_abate_kwon_t'],
                     '탄소 잠재가격 λ (천원/t)': V(c, 'lambda_kwon_t')})
    return pd.DataFrame(rows)


def balance_table():
    its = [('석탄 발전', 'coal_TWh'), ('LNG 발전', 'LNG_power_TWh'), ('CHP 발전', 'CHP_power_TWh'),
           ('CHP 열', 'CHP_heat_TWh'), ('원자력', 'nuclear_TWh'), ('PV 발전', 'PV_TWh'), ('풍력 발전', 'WT_TWh'),
           ('출력제한', 'curtail_TWh'), ('연료전지 발전', 'fuelcell_power_TWh'), ('DR', 'DR_TWh'),
           ('히트펌프 전력', 'HP_elec_TWh'), ('히트펌프 열', 'HP_heat_TWh'), ('열보조공급', 'heat_fallback_TWh'),
           ('수전해 전력', 'electrolyser_elec_TWh'), ('EV 충전', 'EV_charge_TWh'), ('V2G 방전', 'V2G_TWh'),
           ('ESS 충전', 'ESS_charge_TWh'), ('전력 slack', 'power_slack_TWh'), ('가스 연료 총량', 'gas_fuel_TWh'),
           ('HVDC 수송량', 'HVDC_TWh'), ('석탄 이용률 (%)', 'coal_CF_%'),
           ('CO2 석탄 (Mt)', 'CO2_coal_Mt'), ('CO2 LNG 발전 (Mt)', 'CO2_LNG_Mt'), ('CO2 CHP 연료 (Mt)', 'CO2_CHPfuel_Mt'),
           ('CO2 합계 (Mt)', 'CO2_Mt')]
    df = pd.DataFrame({'항목 (TWh)': [a for a, _ in its]})
    for c in CASES:
        df[CASE_KR[c]] = [V(c, k) for _, k in its]
    return df


GROUPS = [('석탄 연료비', ['c_opex_coal']), ('가스 연료비', ['c_opex_gas']), ('CHP 운전비', ['c_opex_CHP']),
          ('원자력 연료비', ['c_opex_nuclear']), ('수소·연료전지', ['c_capex_Electrolyser', 'c_capex_Fuelcell', 'c_opex_Fuelcell',
                                                       'c_opex_Electrolyser']),
          ('저장 설비 투자비', ['c_capex_ESS', 'c_capex_H_Storage', 'c_capex_H2_Storage', 'c_capex_ESS_Ch', 'c_capex_ESS_Disch',
                          'c_capex_EV_DSM'])]


def gsum(c, ks):
    return sum(0 if (k not in nat or pd.isna(nat.at[c, k])) else float(nat.at[c, k]) for k in ks)


def cost_table():
    rows = [g for g, _ in GROUPS] + ['기타', '시스템 비용 합계', '(제외) slack·Fallback 운전 페널티', '(제외) slack 설비 투자비']
    df = pd.DataFrame({'비용 항목 (조원/년)': rows})
    for c in CASES:
        vals = [gsum(c, ks) for _, ks in GROUPS]
        df[CASE_KR[c]] = vals + [V(c, 'cost_sys') - sum(vals), V(c, 'cost_sys'), gsum(c, ['c_opex_penalty']),
                                 gsum(c, ['c_capex_slack'])]
    return df


def capacity_table():
    keys = sorted({k for cp in CAPS.values() for k in cp if '|exist' not in k[0]})
    KR = {'PV': '태양광', 'WT': '풍력', 'ESS': 'ESS 저장', 'ESS_Ch': 'ESS 충전', 'ESS_Disch': 'ESS 방전', 'Electrolyser': '수전해',
          'Fuelcell': '연료전지', 'H2_Storage': '수소 저장', 'H_Storage': '열 저장', 'EV_DSM': 'EV DSM', 'slack': '전력 slack',
          'H_Fallback_Gen': '열보조공급', 'LNG_Fuel_Supply': 'CHP 연료공급'}
    rows = []
    for k, u in keys:
        r = {'설비': f"{KR.get(k, k)} ({u})", '입력 기존': CAPS['B0'].get((k + '|exist', u), 0.0)}
        for c in CASES:
            r[CASE_KR[c]] = CAPS[c].get((k, u), 0.0)
        if max(abs(r[CASE_KR[c]] - r['입력 기존']) for c in CASES) > 0.01 or k in ('PV', 'WT'):
            rows.append(r)
    return pd.DataFrame(rows)


def regional_table():
    rows = []
    for r in REG:
        row = {'지역': REGION_KR[r]}
        for key, kr in (('coal_TWh', '석탄'), ('LNG_TWh', 'LNG')):
            for c in CASES:
                row[f'{kr} {CASE_KR[c]}'] = REGD[c][r][key]
        rows.append(row)
    return pd.DataFrame(rows)


TBL = {'표1_요약': summary_table(), '표2_수급': balance_table(), '표3_비용': cost_table(), '표4_설비': capacity_table(),
       '표5_지역별': regional_table()}
with pd.ExcelWriter(os.path.join(OUT, f'CO2_비용민감도_{args.year}_표.xlsx')) as xw:
    for k, v in TBL.items():
        v.to_excel(xw, sheet_name=k, index=False)
    nat.T.to_excel(xw, sheet_name='전국_전체지표')
nat.to_csv(os.path.join(OUT, 'national.csv'), encoding='utf-8-sig')

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


FM = {'시스템 비용 (조원, slack 제외)': '{:,.3f}', '비용 증가 (조원)': '{:,.3f}', 'slack 제외 증가율 (%)': '{:,.2f}',
      '목적함수 실제 증가율 (%)': '{:,.2f}', '탄소 잠재가격 λ (천원/t)': '{:,.1f}',
      '기준 대비 CO2 (Mt)': '{:+,.2f}', '기준 대비 CO2 (%)': '{:+,.1f}', 'CO2 (Mt)': '{:,.2f}', '평균 감축비용 (천원/t)': '{:,.0f}'}
FM2 = {CASE_KR[c]: '{:,.2f}' for c in CASES}
last = ALT[-1]
cap_tbl = capacity_table()
cap_changed = [r['설비'] for _, r in cap_tbl.iterrows()
               if max(abs(r[CASE_KR[c]] - r[CASE_KR['B0']]) for c in ALT) > 0.01 and not r['설비'].startswith(('태양광', '풍력'))]

doc.add_heading(f'{args.year}년 CO2 제약 해제 조건의 비용 상한별 최소 배출 분석', 0)
P('설비 고정 + CO2 배출 상한 해제 기준에서 시스템 비용 증가 허용률(0·1·3·5·10%)별 최소 배출 운영', size=12)
P(f'작성일: {datetime.now():%Y-%m-%d}   |   모형: PyPSA 0.31.1 + linopy + CPLEX Barrier, 8,784시간 단일연도', size=9)

H('개요', 1)
P('기존 분석은 CO2 배출 상한을 지키는 비용 최소화였다. 이번 분석은 CO2 상한을 사실상 없앤 비용최소해를 기준으로, 시스템 비용을 기준보다 '
  '일정 비율까지 더 쓸 수 있을 때 배출량을 얼마나 줄일 수 있는지 본다. '
  f'입력은 현재 interface의 {args.year}년 조건이며, 모든 경우에 같은 조건을 쓴다.')
T(pd.DataFrame([
    ['기준 (비용최소)', f"현재 interface {args.year}년 입력, CO2 상한 {V('B0', 'CO2_cap_Mt'):.1f} Mt(비구속), 설비 고정. "
                    f"GUI 기준 결과({args.ref_dir})와 목적함수 일치"],
    ['+ε% (ε = 0, 1, 3, 5, 10)', 'PyPSA 목적함수 ≤ (1+ε)·기준 목적함수 아래 CO2 배출량이 가장 낮은 해. '
                                 '목적함수에 탄소가격 λ×배출량을 더해 비용 최소화하고, 상한을 넘지 않는 가장 큰 λ의 해를 택함. '
                                 'ε = 0%는 λ = 1천원/t로, 같은 최소비용 해 중 배출량이 가장 낮은 조합'],
], columns=['경우', '조건']), font=8, caption='표 0-1. 분석 경우')

H('핵심 결과', 2)
B([('기준: ', f"시스템 비용 {C0:.3f}조원/년, CO2 {E0:.1f} Mt (석탄 {V('B0', 'coal_TWh'):.1f} TWh, LNG 발전 "
    f"{V('B0', 'LNG_power_TWh'):.1f} TWh). CO2 상한이 없으므로 연료비가 싼 석탄이 최대한 쓰인다(석탄 이용률 {V('B0', 'coal_CF_%'):.0f}%)."),
   ('같은 최소비용 (+0%): ', f"CO2 {V('CO2min-0', 'CO2_Mt'):.2f} Mt ({V('CO2min-0', 'dE_Mt'):+.2f} Mt). "
    + ('비용최소해가 사실상 유일해서 배출량이 거의 같다.' if abs(V('CO2min-0', 'dE_Mt')) < 1.0 else
       '비용이 같은 해 중에서도 배출량을 줄일 여지가 있다.')),
   ('비용 증가 허용: ', ', '.join(f"+{EPS[c]:g}% → {V(c, 'CO2_Mt'):.1f} Mt ({V(c, 'dE_%'):+.1f}%)" for c in ALT if c != 'CO2min-0') + '.'),
   ('감축 수단: ', f"{CASE_KR[last]}에서 석탄 발전 {d(last, 'coal_TWh'):+.1f} TWh, LNG 발전 {d(last, 'LNG_power_TWh'):+.1f} TWh, "
    f"CHP 발전 {d(last, 'CHP_power_TWh'):+.1f} TWh, 출력제한 {d(last, 'curtail_TWh'):+.1f} TWh. 설비가 고정되어 있으므로 "
    '감축은 석탄 → LNG 연료 전환(운전 순서 변경)이 중심이다.'),
   ('감축비용: ', ', '.join(f"{CASE_KR[c]} {nat.at[c, 'avg_abate_kwon_t']:,.0f}" for c in ALT if pd.notna(nat.at[c, 'avg_abate_kwon_t']))
    + ' 천원/tCO2 (기준 대비 평균). 구간 한계 감축비용은 그림 4 참고.')])
T(summary_table(), font=7, caption='표 0-2. 결과 요약', fmts=FM)
F('fig1_pareto.png', '그림 1. 비용 증가율별 최소 배출량(a)과 배출원별 CO2(b)')

doc.add_page_break()
H('1. 수급 변화', 1)
F('fig2_mix.png', '그림 2. 발전원별 발전량(a)과 기준 대비 연간 수급 변화(b)', width=17)
T(balance_table(), font=7, caption='표 1-1. 전력·열·가스 수급 (TWh)', fmts=FM2)
F('fig3_regional.png', '그림 3. 지역별 석탄(a)·LNG(b) 발전량', width=17)
T(regional_table(), font=6, caption='표 1-2. 지역별 석탄·LNG 발전량 (TWh)', fmt='{:,.1f}')

doc.add_page_break()
H('2. 비용과 감축비용', 1)
T(cost_table(), font=7.5, caption='표 2-1. 비용 구성 (조원/년)', fmts={k: '{:,.3f}' for k in FM2})
F('fig4_abatement.png', '그림 4. CO2 감축비용 (평균: 기준 대비, 구간 한계: 직전 단계 대비)', width=14)
B([('읽는 법: ', '평균 감축비용 = (해당 경우 비용 − 기준 비용) ÷ (기준 배출 − 해당 배출). 구간 한계 감축비용은 바로 앞 단계 대비 추가 1톤을 줄이는 데 든 비용이다.'),
   ('비교: ', f"기존 CO2 상한 분석의 CO2 잠재가격(shadow price)과 같은 개념이다. 연료 전환 비용은 석탄·LNG 연료비와 효율 차이로 정해진다.")])

H('3. 설비', 1)
P('설비 용량은 기준과 같은 입력(확장 가능 설비의 상한 포함)을 쓰며, 대안 단계에서도 기준과 같은 조건이다. '
  + (f"대안에서 용량이 달라진 설비: {', '.join(cap_changed)}." if cap_changed else '대안에서 기준과 용량이 달라진 설비는 없다.'))
T(cap_tbl, font=7, caption='표 3-1. 설비 용량 (GW / GWh)', fmt='{:,.2f}')

H('부록. 방법·실행 정보', 1)
B([('비용 상한: ', '최적화에서는 기존 분석과 같은 PyPSA 목적함수(slack·Fallback 페널티 포함)를 그대로 쓰고, 비용 상한도 이 목적함수 기준이다.'),
   ('시스템 비용 (보고용): ', '목적함수에서 slack 비용을 후처리로 모두 뺀 값. '
    f"제외 항목은 전력 slack·EV failsafe·열/수소 보조공급(Fallback)의 운전 페널티(기준 {V('B0', 'penalty_cost') / 1e9:.2f}조원)와 "
    f"전력 slack 설비 투자비(기준 {V('B0', 'slack_capex') / 1e9:.3f}조원)이다."),
   ('탐색 방법: ', '비용 상한 제약 + 배출량 목적함수 형태는 CPLEX Barrier가 수렴하지 못해, 같은 효율 경계를 탄소가격 가중합(목적함수 + λ·배출량)으로 구했다. '
    'LP에서는 두 방법이 같은 효율 경계를 준다. 효율 경계가 꺾이는 점 사이에서는 상한을 정확히 채우는 λ가 없을 수 있어, 실제 증가율은 목표 이하의 가장 가까운 값이다(표 0-2).'),
   ('보조공급: ', '보조공급 단가(약 500만원/MWh)가 탐색한 λ보다 훨씬 커서 보조공급이 감축 수단으로 쓰이지 않는다(표 1-1의 열보조공급·전력 slack 참고).'),
   ('배출량: ', '발전기 연료 사용 × 배출계수(석탄 0.359, 가스 0.211 tCO2/MWh_fuel). CHP 연료는 CHP 연료공급 발전기(LNG_Fuel_Supply)로 포함된다. PyPSA CO2_limit 제약과 같은 산식이다.'),
   ('기준 검증: ', f"다시 계산한 기준 목적함수 {V('B0', 'objective_full') / 1e9:.5f}조원 = GUI 결과와 일치."),
   ('실행: ', f"run_co2_price_search.py → {args.run_dir} (λ별 결과 lambda_log.csv), 보고서 make_co2cost_report.py → {OUT}")])

path = os.path.join(OUT, f'CO2_비용민감도_{args.year}_분석_보고서.docx')
doc.save(path)
print(f'완료: {path}')
pd.set_option('display.width', 250); pd.set_option('display.max_columns', 30)
print(summary_table().round(3).to_string(index=False))
print(balance_table().round(2).to_string(index=False))
print(cost_table().round(3).to_string(index=False))
print(cap_tbl.round(2).to_string(index=False))
