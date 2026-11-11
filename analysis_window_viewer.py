"""
Analysis Window 뷰어 — 기존 결과 엑셀을 재분석 없이 시간별로 점검

사용법
------
  GUI   : python analysis_window_viewer.py [결과폴더]
  배치  : python analysis_window_viewer.py --file 결과.xlsx [--compare 결과2.xlsx]
                 --start "2024-08-01 00:00" --hours 168
                 [--auto line|curtail|demand|net_low] [--png 그림.png] [--export AW_결과.xlsx]

폴더 안의 결과 파일(*.xlsx) 중 하나(또는 비교용 두 개)를 골라 지정 기간의
  · 전력수급 그림 (Figure_hour_electricity 형식: 공급 누적 + 수요선 / 유연 소비 누적)
  · Analysis_Window 7개 섹션 (송전 회랑, 발전원, 링크, 저장량, 저장 출입, Curtailment, 수요)
을 보여주고, 현재 구간을 엑셀(차트 포함)로 내보낸다.
탭마다 표시 항목·쌓는 순서·범례 위치·그래프 형태를 바꿀 수 있고, 설정은 다음 실행에도 유지된다.
읽은 결과는 결과 파일 옆 _aw_cache 폴더에 캐시되어 다음부터 즉시 열린다.
"""
import argparse
import json
import os
import queue
import sys
import threading

import pandas as pd

import analysis_window_core as core

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS = os.path.join(BASE_DIR, 'analysis_window_viewer.json')
PERIODS = {'1일': 24, '3일': 72, '7일': 168, '14일': 336, '30일': 720}
AUTO_MODES = {'선로 포화 최대': 'line', 'Curtailment 최대': 'curtail',
              '전력수요 최대': 'demand', '순부하 최소(재생 최대)': 'net_low'}
SEC_COLORS = {'corridor': '1F4E79', 'gen': '375623', 'link': '7030A0', 'store_e': 'C55A11',
              'store_p': '833C00', 'curtail': '843C0C', 'loads': '155F8A'}
TABS = [('balance', '전력수급 (Figure_hour_electricity)')] + core.SECTIONS
LEGEND_POS = ['오른쪽 바깥', '아래 바깥', '그래프 안 - 자동', '그래프 안 - 왼쪽 위', '그래프 안 - 오른쪽 위', '숨김']
KINDS = ['누적 면적', '선']
DEFAULT_KIND = {'balance': '누적 면적', 'gen': '누적 면적', 'curtail': '누적 면적'}
NO_FILE = '(비교 안 함)'


# ── 탭별 계열 정의 ───────────────────────────────────────────────────
def tab_items(obj, key):
    """탭에 그릴 수 있는 계열 목록.
    각 항목: id(파일 간 공통 키), label, row(전력수급: 0=공급, 1=소비), fixed_line, series, color, on(기본 표시)."""
    items = []

    def add(id_, label, series, color=None, on=True, row=0, fixed_line=False, style=None):
        items.append(dict(id=id_, label=label, series=series, color=color, on=bool(on), row=row,
                          fixed_line=fixed_line, style=style or {}))

    if key == 'balance':
        b = obj['balance']
        lab_s = {k: (l, c) for k, l, c in core.SUPPLY_DEF}
        lab_c = {k: (l, c) for k, l, c in core.CONSUME_DEF}
        for k in b['supply']:
            s = b['supply'][k].clip(lower=0)
            l, c = lab_s.get(k, (k, None))
            add('S:' + k, l, s, c, s.max() > 1e-3, 0)
        flex = b['consume'].drop(columns=['Curtail'], errors='ignore').sum(axis=1)
        add('L:demand', core.DEMAND_LABEL, b['demand'], '#000000', True, 0, True, {'lw': 1.6})
        add('L:demand_flex', '전력수요+유연소비', b['demand'] + flex, '#000000', True, 0, True, {'lw': 1.0, 'ls': '--'})
        for k in b['consume']:
            s = b['consume'][k].clip(lower=0)
            l, c = lab_c.get(k, (k, None))
            add('C:' + k, l, s, c, s.max() > 1e-3, 1)
        return items

    sec = obj['sections'].get(key)
    if sec is None:
        return items
    if key == 'curtail':
        add('태양광 합계', '태양광 합계', sec[[c for c in sec if c.endswith('_PV')]].sum(axis=1), '#FFD966')
        add('풍력 합계', '풍력 합계', sec[[c for c in sec if c.endswith('_WT')]].sum(axis=1), '#70AD47')
        for c in sec:
            add(c, c, sec[c], on=False)
    elif key == 'gen':
        el = set(obj['balance']['supply'].columns)
        lab = {k: (l, c) for k, l, c in core.SUPPLY_DEF}
        for c in sec:
            l, col = lab.get(c, (c, None))
            add(c, l, sec[c], col, c in el and sec[c].abs().max() > 1e-3)
    else:
        for c in sec:
            add(c, c, sec[c], on=sec[c].abs().max() > 1e-3)
    return items


def item_display(key, it):
    if key != 'balance':
        return it['label']
    return {'S': '[공급] ', 'L': '[선] ', 'C': '[소비] '}[it['id'][0]] + it['label']


def merge_state(key, state, objs):
    """여러 파일의 계열을 합쳐 순서 목록을 갱신(기존 순서 유지, 새 항목은 뒤에)."""
    st = state.setdefault(key, {})
    st.setdefault('order', [])
    st.setdefault('checked', {})
    st.setdefault('legend', LEGEND_POS[0])
    st.setdefault('kind', DEFAULT_KIND.get(key, '선'))
    labels, defaults = {}, {}
    for obj in objs:
        for it in tab_items(obj, key):
            if it['id'] not in labels:
                labels[it['id']] = item_display(key, it)
                defaults[it['id']] = it['on']
            else:
                defaults[it['id']] = defaults[it['id']] or it['on']
    st['order'] = [i for i in st['order'] if i in labels] + [i for i in labels if i not in st['order']]
    st['_labels'] = labels
    st['_defaults'] = defaults
    st['_default_order'] = list(labels)
    return st


def is_checked(st, id_):
    return st['checked'].get(id_, st.get('_defaults', {}).get(id_, True))


# ── 그림 ─────────────────────────────────────────────────────────────
def _mpl_setup():
    import matplotlib
    matplotlib.rcParams['font.family'] = 'Malgun Gothic'
    matplotlib.rcParams['axes.unicode_minus'] = False


def _fmt_time_axis(ax, idx):
    import matplotlib.dates as mdates
    import matplotlib.ticker as mticker
    hours = len(idx)
    if hours <= 72:
        loc = mdates.HourLocator(byhour=[0, 6, 12, 18])
    elif hours <= 336:
        loc = mdates.DayLocator()
    else:
        loc = mdates.DayLocator(interval=3)
    ax.xaxis.set_major_locator(loc)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%m-%d\n%H시'))
    ax.set_xlim(idx[0], idx[-1])
    ax.grid(alpha=0.3)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'{v:,.0f}'))


def _place_legend(row_axes, handles, labels, pos, n_panes):
    if pos == '숨김' or not handles:
        return
    kw = dict(fontsize=8, frameon=False)
    if pos == '오른쪽 바깥':
        row_axes[-1].legend(handles, labels, loc='upper left', bbox_to_anchor=(1.01, 1.0),
                            ncol=2 if len(handles) > 16 else 1, **kw)
    elif pos == '아래 바깥':
        ncol = min(len(handles), 7 if n_panes == 1 else 4)
        for ax in row_axes:
            ax.legend(handles, labels, loc='upper center', bbox_to_anchor=(0.5, -0.12), ncol=ncol, **kw)
    else:
        loc = {'그래프 안 - 자동': 'best', '그래프 안 - 왼쪽 위': 'upper left',
               '그래프 안 - 오른쪽 위': 'upper right'}[pos]
        for ax in row_axes:
            ax.legend(handles, labels, loc=loc, ncol=2 if len(handles) > 12 else 1,
                      fontsize=8, frameon=True, framealpha=0.85)


def plot_tab(fig, key, panes, st, sharey=True):
    """panes: [(제목, obj, slice)] — 1개면 단일, 2개면 좌/우 비교."""
    import matplotlib
    fig.clear()
    if not panes:
        return
    nrows = 2 if key == 'balance' else 1
    ncols = len(panes)
    axes = fig.subplots(nrows, ncols, squeeze=False, sharex='col',
                        sharey='row' if (sharey and ncols > 1) else 'none')
    cmap = matplotlib.colormaps['tab20']
    order = st['order']
    stack_mode = key == 'balance' or st.get('kind') == '누적 면적'
    unit = 'MWh' if key in ('store_e', 'curtail') else 'MW'
    row_label = ['공급 (MW)', '유연 소비 (MW)'] if key == 'balance' else [unit]
    # 행별 범례 핸들: 쌓기(아래→위 순서)와 선 분리
    leg = [{'stack': {}, 'line': {}} for _ in range(nrows)]

    def color_of(it):
        if it['color']:
            return it['color']
        return cmap(order.index(it['id']) % 20 if it['id'] in order else 0)

    for c, (title, obj, sl) in enumerate(panes):
        items = {it['id']: it for it in tab_items(obj, key)}
        ids = [i for i in order if i in items and is_checked(st, i)]
        idx = obj['index'][sl]
        for r in range(nrows):
            ax = axes[r][c]
            if not items:
                ax.text(0.5, 0.5, '원본 시트가 없어 표시할 수 없습니다', ha='center', va='center',
                        transform=ax.transAxes)
                continue
            rit = [items[i] for i in ids if items[i]['row'] == r]
            st_items = [it for it in rit if stack_mode and not it['fixed_line']]
            ln_items = [it for it in rit if it not in st_items]
            if st_items:
                polys = ax.stackplot(idx, *[it['series'].iloc[sl].values for it in st_items],
                                     colors=[color_of(it) for it in st_items], alpha=0.9, linewidth=0)
                for it, p in zip(st_items, polys):
                    leg[r]['stack'].setdefault(it['label'], p)
            for it in ln_items:
                style = it['style'] or {'lw': 1.2}
                h, = ax.plot(idx, it['series'].iloc[sl].values, color=color_of(it), **style)
                leg[r]['line'].setdefault(it['label'], h)
            if not stack_mode or key in ('corridor', 'store_p', 'link'):
                ax.axhline(0, color='black', lw=0.7)
            _fmt_time_axis(ax, idx)
            if r == 0:
                ax.set_title(title, loc='left', fontsize=10)
            if c == 0:
                ax.set_ylabel(row_label[r])
    pos = st.get('legend', LEGEND_POS[0])
    for r in range(nrows):
        lines = list(leg[r]['line'].items())
        stacks = list(leg[r]['stack'].items())
        if pos != '아래 바깥':
            stacks = stacks[::-1]   # 세로 범례: 위쪽 항목 = 그래프 위쪽
        entries = lines + stacks
        _place_legend(axes[r], [h for _, h in entries], [l for l, _ in entries], pos, ncols)


def tab_table(obj, key, st, sl):
    items = {it['id']: it for it in tab_items(obj, key)}
    ids = [i for i in st['order'] if i in items and is_checked(st, i)]
    df = pd.DataFrame({item_display(key, items[i]): items[i]['series'].iloc[sl] for i in ids},
                      index=obj['index'][sl])
    if key == 'balance':
        df['수급차(공급-소비)'] = obj['balance']['residual'].iloc[sl]
    return df


def pane_title(obj, tag=None):
    name = os.path.splitext(os.path.basename(obj['path']))[0]
    return f'[{tag}] {name}' if tag else name


# ── 엑셀 내보내기 ────────────────────────────────────────────────────
def export_excel(obj, sl, out_path, balance_state=None):
    """balance_state 를 주면 전력수급 차트의 계열 선택·쌓는 순서를 화면과 같게 맞춘다."""
    from openpyxl import Workbook
    from openpyxl.chart import AreaChart, LineChart, Reference
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    bt = core.balance_table(obj, sl)
    idx = bt.index
    wb = Workbook()

    info = wb.active
    info.title = 'Info'
    rows = [('원본 결과 파일', obj['path']),
            ('분석 시작', idx[0].strftime('%Y-%m-%d %H:%M')),
            ('분석 종료', idx[-1].strftime('%Y-%m-%d %H:%M')),
            ('시간 수', len(idx)),
            ('송전 회랑 데이터', obj['sections'].get('_corridor_src', '-')),
            ('수급 검증 |공급-소비| 최대 (MW)', float(bt['수급차(공급-소비)'].abs().max())),
            ('비고', '전력수급은 전력모선(EL) 기준. 수요 = 일반 전력수요 + 계통 공급 DC/Fab, '
                   '유연 소비 = 히트펌프·수전해·EV 충전·ESS/양수 충전·HVDC 손실 등. '
                   'Curtailment 는 참고값(수급 합계에서 제외).')]
    for r, (k, v) in enumerate(rows, start=1):
        info.cell(r, 1, k).font = Font(bold=True)
        info.cell(r, 2, v)
    info.column_dimensions['A'].width = 30
    info.column_dimensions['B'].width = 100

    # Figure_hour_electricity: 표 + 차트
    ws = wb.create_sheet('Figure_hour_electricity')
    sup_cols = [c for c in bt.columns if c.startswith('[공급]')]
    dem_col = f'[소비] {core.DEMAND_LABEL}'
    con_cols = [c for c in bt.columns if c.startswith('[소비]') and c != dem_col]
    col_order = ['시각'] + sup_cols + [dem_col] + con_cols + ['수급차(공급-소비)']
    for j, c in enumerate(col_order, start=1):
        cell = ws.cell(1, j, c.replace('[공급] ', '').replace('[소비] ', ''))
        cell.font = Font(bold=True, size=9)
        cell.fill = PatternFill('solid', fgColor='DDEBF7' if c in sup_cols else 'FCE4D6' if c in con_cols else 'EDEDED')
        cell.alignment = Alignment(wrap_text=True, vertical='center')
    for i, ts in enumerate(idx, start=2):
        ws.cell(i, 1, ts.strftime('%m-%d %H시'))
        for j, c in enumerate(col_order[1:], start=2):
            ws.cell(i, j, round(float(bt.at[ts, c]), 3)).number_format = '#,##0'
    ws.freeze_panes = 'B2'
    ws.row_dimensions[1].height = 42
    n = len(idx)
    cats = Reference(ws, min_col=1, min_row=2, max_row=n + 1)
    colors = {f'[공급] {l}': c for _, l, c in core.SUPPLY_DEF}
    colors.update({f'[소비] {l}': c for _, l, c in core.CONSUME_DEF})

    chart_sup, chart_con, show_demand = sup_cols, con_cols, True
    if balance_state:
        lab_s = {k: l for k, l, _ in core.SUPPLY_DEF}
        lab_c = {k: l for k, l, _ in core.CONSUME_DEF}
        chosen = [i for i in balance_state['order'] if is_checked(balance_state, i)]
        chart_sup = [f'[공급] {lab_s.get(i[2:], i[2:])}' for i in chosen if i.startswith('S:')]
        chart_con = [f'[소비] {lab_c.get(i[2:], i[2:])}' for i in chosen if i.startswith('C:')]
        chart_sup = [c for c in chart_sup if c in sup_cols]
        chart_con = [c for c in chart_con if c in con_cols]
        show_demand = 'L:demand' in chosen

    def _area(cols, title):
        ch = AreaChart()
        ch.grouping = 'stacked'
        ch.title = title
        ch.height, ch.width = 11, 30
        for c in cols:
            if bt[c].abs().max() <= 1e-3:
                continue
            j = col_order.index(c) + 1
            ch.add_data(Reference(ws, min_col=j, min_row=1, max_row=n + 1), titles_from_data=True)
            s = ch.series[-1]
            hexc = colors.get(c, '#999999').lstrip('#')
            s.graphicalProperties.solidFill = hexc
            s.graphicalProperties.line.noFill = True
        ch.set_categories(cats)
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.x_axis.tickLblSkip = 24 if n > 72 else 6
        ch.x_axis.tickMarkSkip = 24 if n > 72 else 6
        ch.y_axis.numFmt = '#,##0'
        ch.y_axis.majorGridlines = None
        ch.legend.position = 'b'
        ch.legend.overlay = False
        return ch

    ch1 = _area(chart_sup, '시간별 전력수급 — 공급(누적)과 전력수요 (MW)')
    if show_demand:
        ln = LineChart()
        j = col_order.index(dem_col) + 1
        ln.add_data(Reference(ws, min_col=j, min_row=1, max_row=n + 1), titles_from_data=True)
        ln.series[0].graphicalProperties.line.solidFill = '000000'
        ln.series[0].graphicalProperties.line.width = 22000
        ln.series[0].smooth = False
        ch1 += ln
    ch2 = _area(chart_con, '유연 소비(누적) 및 Curtailment (MW)')
    anchor_col = get_column_letter(len(col_order) + 2)
    ws.add_chart(ch1, f'{anchor_col}2')
    ws.add_chart(ch2, f'{anchor_col}25')

    # Analysis_Window: 7개 섹션 (값)
    aw = wb.create_sheet('Analysis_Window')
    aw.cell(1, 1, '■ Analysis Window – 섹터커플링 운영 현황 분석').font = Font(bold=True, size=13, color='1F4E79')
    aw.cell(2, 1, f'원본: {os.path.basename(obj["path"])}   기간: {idx[0]:%Y-%m-%d %H:%M} ~ {idx[-1]:%Y-%m-%d %H:%M} '
                  f'({n}h)').font = Font(italic=True, size=9)
    row = 4
    for key, title in core.SECTIONS:
        df = obj['sections'].get(key)
        if df is None:
            continue
        df = df.iloc[sl]
        fill = PatternFill('solid', fgColor=SEC_COLORS[key])
        c = aw.cell(row, 1, title if key != 'corridor' else f'{title}  [{obj["sections"].get("_corridor_src")}]')
        c.font = Font(bold=True, color='FFFFFF', size=10)
        for ci in range(1, len(df.columns) + 2):
            aw.cell(row, ci).fill = fill
        row += 1
        aw.cell(row, 1, 'Timestamp').font = Font(bold=True, size=9)
        for ci, cname in enumerate(df.columns, start=2):
            aw.cell(row, ci, cname).font = Font(bold=True, size=9)
            aw.cell(row, ci).fill = PatternFill('solid', fgColor='BDD7EE')
        aw.cell(row, 1).fill = PatternFill('solid', fgColor='BDD7EE')
        row += 1
        for ts, vals in zip(df.index, df.itertuples(index=False)):
            aw.cell(row, 1, ts.to_pydatetime()).number_format = 'YYYY-MM-DD HH:MM'
            for ci, v in enumerate(vals, start=2):
                aw.cell(row, ci, round(float(v), 3)).number_format = '#,##0.0'
            row += 1
        row += 1
    aw.column_dimensions['A'].width = 18
    wb.save(out_path)
    return out_path


# ── 배치 실행 ────────────────────────────────────────────────────────
def run_batch(a):
    objs = [core.load_result(a.file, use_cache=not a.no_cache, progress=print)]
    if a.compare:
        objs.append(core.load_result(a.compare, use_cache=not a.no_cache, progress=print))
    start = core.auto_start(objs[0], a.auto, a.hours) if a.auto else (a.start or objs[0]['index'][0])
    sls = [core.window(o, start, a.hours) for o in objs]
    for o, sl in zip(objs, sls):
        idx = o['index'][sl]
        print(f'{os.path.basename(o["path"])}: {idx[0]} ~ {idx[-1]} ({len(idx)}h), '
              f'회랑 데이터: {o["sections"].get("_corridor_src")}, '
              f'수급 검증 최대오차 {o["balance"]["residual"].iloc[sl].abs().max():.2e} MW')
    if a.png:
        import matplotlib
        matplotlib.use('Agg')
        _mpl_setup()
        import matplotlib.pyplot as plt
        tags = ['A', 'B'] if len(objs) > 1 else [None]
        panes = [(pane_title(o, t), o, sl) for o, sl, t in zip(objs, sls, tags)]
        state = {}
        for key, _ in TABS if a.all_png else TABS[:1]:
            st = merge_state(key, state, objs)
            st['legend'] = a.legend
            fig = plt.figure(figsize=(9 * len(objs) + 6, 8 if key == 'balance' else 5.5), layout='constrained')
            plot_tab(fig, key, panes, st)
            p = a.png if key == 'balance' else os.path.splitext(a.png)[0] + f'_{key}.png'
            fig.savefig(p, dpi=110)
            plt.close(fig)
            print('그림 저장:', p)
    if a.export:
        print('엑셀 저장:', export_excel(objs[0], sls[0], a.export))


# ── GUI: 계열 선택 패널 ──────────────────────────────────────────────
class SeriesPanel:
    """체크(표시)·순서(위=먼저 쌓임)·검색·범례 위치·그래프 형태 조정."""
    CHECK, UNCHECK = '☑', '☐'

    def __init__(self, master, key, app):
        import tkinter as tk
        from tkinter import ttk
        self.key, self.app = key, app
        self.frame = ttk.Frame(master, padding=(4, 2))
        f = self.frame
        ttk.Label(f, text='표시 항목  (목록 위 = 아래부터 쌓임)', font=('Malgun Gothic', 9, 'bold')).pack(anchor='w')
        sr = ttk.Frame(f)
        sr.pack(fill='x', pady=2)
        self.var_q = tk.StringVar()
        e = ttk.Entry(sr, textvariable=self.var_q, width=12)
        e.pack(side='left', fill='x', expand=True)
        e.bind('<Return>', lambda ev: self.check_match(only=True))
        ttk.Button(sr, text='일치만', width=6, command=lambda: self.check_match(only=True)).pack(side='left', padx=1)
        ttk.Button(sr, text='+추가', width=5, command=lambda: self.check_match(only=False)).pack(side='left')
        ttk.Label(f, text='검색어 여러 개는 쉼표(,)로 구분. 예: GND, ICN', foreground='#777777',
                  font=('Malgun Gothic', 8)).pack(anchor='w')

        tf = ttk.Frame(f)
        tf.pack(fill='both', expand=True, pady=2)
        self.tree = ttk.Treeview(tf, show='tree', selectmode='extended', height=18)
        self.tree.column('#0', width=210)
        ys = ttk.Scrollbar(tf, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=ys.set)
        ys.pack(side='right', fill='y')
        self.tree.pack(side='left', fill='both', expand=True)
        self.tree.bind('<Button-1>', self._on_click)
        self.tree.bind('<Double-1>', lambda ev: 'break')
        self.tree.bind('<space>', lambda ev: (self.toggle(self.tree.selection()), 'break')[1])

        br = ttk.Frame(f)
        br.pack(fill='x')
        for txt, cmd in [('▲', lambda: self.move(-1)), ('▼', lambda: self.move(1)),
                         ('맨 위', lambda: self.move_end(True)), ('맨 아래', lambda: self.move_end(False))]:
            ttk.Button(br, text=txt, width=5 if len(txt) > 1 else 3, command=cmd).pack(side='left', padx=1)
        br2 = ttk.Frame(f)
        br2.pack(fill='x', pady=2)
        for txt, cmd in [('전체 표시', lambda: self.check_all(True)), ('전체 해제', lambda: self.check_all(False)),
                         ('기본값', self.reset)]:
            ttk.Button(br2, text=txt, command=cmd).pack(side='left', padx=1, fill='x', expand=True)
        ttk.Label(f, text='☐ 클릭: 표시 전환 / 이름 클릭: 선택(Shift·Ctrl 다중) → ▲▼ 이동',
                  foreground='#777777', font=('Malgun Gothic', 8)).pack(anchor='w')

        og = ttk.Frame(f)
        og.pack(fill='x', pady=(6, 0))
        ttk.Label(og, text='범례 위치').grid(row=0, column=0, sticky='w')
        self.var_leg = tk.StringVar()
        cb = ttk.Combobox(og, textvariable=self.var_leg, values=LEGEND_POS, state='readonly', width=18)
        cb.grid(row=0, column=1, sticky='ew', padx=2, pady=1)
        cb.bind('<<ComboboxSelected>>', lambda ev: self._set('legend', self.var_leg.get()))
        self.var_kind = tk.StringVar()
        if key != 'balance':
            ttk.Label(og, text='그래프 형태').grid(row=1, column=0, sticky='w')
            cb2 = ttk.Combobox(og, textvariable=self.var_kind, values=KINDS, state='readonly', width=18)
            cb2.grid(row=1, column=1, sticky='ew', padx=2, pady=1)
            cb2.bind('<<ComboboxSelected>>', lambda ev: self._set('kind', self.var_kind.get()))
        og.columnconfigure(1, weight=1)

    @property
    def st(self):
        return self.app.state[self.key]

    def refresh(self):
        st = self.st
        sel = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        for i in st['order']:
            mark = self.CHECK if is_checked(st, i) else self.UNCHECK
            self.tree.insert('', 'end', iid=i, text=f'{mark} {st["_labels"].get(i, i)}')
        keep = [s for s in sel if self.tree.exists(s)]
        if keep:
            self.tree.selection_set(keep)
            self.tree.see(keep[0])
        self.var_leg.set(st['legend'])
        self.var_kind.set(st['kind'])

    def _changed(self):
        self.refresh()
        self.app.redraw(self.key)

    def _set(self, k, v):
        self.st[k] = v
        self.app.redraw(self.key)

    def _on_click(self, ev):
        iid = self.tree.identify_row(ev.y)
        if iid and ev.x < 36:
            self.toggle([iid])
            return 'break'
        return None

    def toggle(self, ids):
        st = self.st
        for i in ids:
            st['checked'][i] = not is_checked(st, i)
        self._changed()

    def check_all(self, on):
        st = self.st
        for i in st['order']:
            st['checked'][i] = on
        self._changed()

    def check_match(self, only):
        terms = [t.strip().lower() for t in self.var_q.get().split(',') if t.strip()]
        if not terms:
            return
        st = self.st
        for i in st['order']:
            hit = any(t in (st['_labels'].get(i, i) + ' ' + i).lower() for t in terms)
            if hit:
                st['checked'][i] = True
            elif only:
                st['checked'][i] = False
        self._changed()

    def move(self, d):
        st = self.st
        sel = [i for i in st['order'] if i in set(self.tree.selection())]
        if not sel:
            return
        order = st['order']
        seq = sel if d < 0 else sel[::-1]
        for i in seq:
            p = order.index(i)
            q = p + d
            if 0 <= q < len(order) and order[q] not in sel:
                order[p], order[q] = order[q], order[p]
        self._changed()

    def move_end(self, top):
        st = self.st
        selset = set(self.tree.selection())
        sel = [i for i in st['order'] if i in selset]
        if not sel:
            return
        rest = [i for i in st['order'] if i not in selset]
        st['order'] = sel + rest if top else rest + sel
        self._changed()

    def reset(self):
        st = self.st
        st['order'] = list(st['_default_order'])
        st['checked'] = {}
        st['legend'] = LEGEND_POS[0]
        st['kind'] = DEFAULT_KIND.get(self.key, '선')
        self._changed()


# ── GUI 본체 ─────────────────────────────────────────────────────────
class ViewerApp:
    def __init__(self, root, folder=None):
        import tkinter as tk
        from tkinter import ttk
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
        from matplotlib.figure import Figure
        _mpl_setup()
        self.root = root
        self.objs = {'A': None, 'B': None}
        self.sls = {}
        self.files = []
        self.q = queue.Queue()
        self.dirty = set()
        self._redraw_job = None
        cfg = self._load_cfg()
        self.state = cfg.get('tabs', {})
        for key, _ in TABS:
            st = self.state.setdefault(key, {})
            st.setdefault('order', [])
            st.setdefault('checked', {})
            st.setdefault('legend', LEGEND_POS[0])
            st.setdefault('kind', DEFAULT_KIND.get(key, '선'))
            st.setdefault('_labels', {})
            st.setdefault('_defaults', {})
            st.setdefault('_default_order', [])

        root.title('Analysis Window 뷰어 — PyPSA 결과 시간별 점검')
        root.geometry('1700x980')

        top = ttk.Frame(root, padding=6)
        top.pack(fill='x')
        r1 = ttk.Frame(top)
        r1.pack(fill='x', pady=2)
        ttk.Label(r1, text='결과 폴더', width=9).pack(side='left')
        self.var_folder = tk.StringVar(value=folder or cfg.get('folder') or BASE_DIR)
        ttk.Entry(r1, textvariable=self.var_folder, width=110).pack(side='left', padx=4, fill='x', expand=True)
        ttk.Button(r1, text='폴더 선택', command=self._pick_folder).pack(side='left')
        self.var_rec = tk.BooleanVar(value=cfg.get('recursive', True))
        ttk.Checkbutton(r1, text='하위 폴더 포함', variable=self.var_rec, command=self._refresh_files).pack(side='left', padx=6)
        ttk.Button(r1, text='새로고침', command=self._refresh_files).pack(side='left')

        self.var_file, self.cb_file = {}, {}
        for tag, label in [('A', '파일 A'), ('B', '파일 B (비교)')]:
            r = ttk.Frame(top)
            r.pack(fill='x', pady=1)
            ttk.Label(r, text=label, width=12).pack(side='left')
            self.var_file[tag] = tk.StringVar()
            cb = ttk.Combobox(r, textvariable=self.var_file[tag], width=120, state='readonly')
            cb.pack(side='left', padx=4, fill='x', expand=True)
            cb.bind('<<ComboboxSelected>>', lambda e, t=tag: self._load(t))
            self.cb_file[tag] = cb
            ttk.Button(r, text='캐시 무시 재읽기', command=lambda t=tag: self._load(t, no_cache=True)).pack(side='left', padx=2)
            if tag == 'B':
                ttk.Button(r, text='A ↔ B 바꾸기', command=self._swap).pack(side='left', padx=2)

        r3 = ttk.Frame(top)
        r3.pack(fill='x', pady=3)
        ttk.Label(r3, text='분석 시작').pack(side='left')
        self.var_start = tk.StringVar()
        e = ttk.Entry(r3, textvariable=self.var_start, width=17)
        e.pack(side='left', padx=4)
        e.bind('<Return>', lambda ev: self._apply())
        ttk.Label(r3, text='기간').pack(side='left', padx=(8, 0))
        self.var_period = tk.StringVar(value=cfg.get('period', '7일'))
        cb = ttk.Combobox(r3, textvariable=self.var_period, values=list(PERIODS), width=6, state='readonly')
        cb.pack(side='left', padx=4)
        cb.bind('<<ComboboxSelected>>', lambda ev: self._apply())
        ttk.Label(r3, text='주차').pack(side='left', padx=(8, 0))
        self.var_week = tk.IntVar(value=1)
        sp = ttk.Spinbox(r3, from_=1, to=53, textvariable=self.var_week, width=4, command=self._apply_week)
        sp.pack(side='left', padx=4)
        sp.bind('<Return>', lambda ev: self._apply_week())
        ttk.Button(r3, text='◀ 이전', command=lambda: self._shift(-1)).pack(side='left', padx=(8, 0))
        ttk.Button(r3, text='다음 ▶', command=lambda: self._shift(1)).pack(side='left', padx=2)
        ttk.Button(r3, text='적용', command=self._apply).pack(side='left', padx=6)
        ttk.Label(r3, text='자동 구간(A 기준)').pack(side='left', padx=(12, 0))
        self.var_auto = tk.StringVar(value='선로 포화 최대')
        ttk.Combobox(r3, textvariable=self.var_auto, values=list(AUTO_MODES), width=20,
                     state='readonly').pack(side='left', padx=4)
        ttk.Button(r3, text='찾기', command=self._auto).pack(side='left')
        self.var_sharey = tk.BooleanVar(value=cfg.get('sharey', True))
        ttk.Checkbutton(r3, text='비교 시 Y축 동일', variable=self.var_sharey,
                        command=lambda: self.redraw_all()).pack(side='left', padx=12)
        ttk.Button(r3, text='엑셀 내보내기', command=self._export).pack(side='right')
        ttk.Button(r3, text='그림 저장(PNG)', command=self._save_png).pack(side='right', padx=4)

        self.var_status = tk.StringVar(value='결과 파일을 선택하세요.')
        ttk.Label(root, textvariable=self.var_status, anchor='w', relief='sunken', padding=3).pack(side='bottom', fill='x')

        self.nb = ttk.Notebook(root)
        self.nb.pack(fill='both', expand=True)
        self.tabs = {}
        for key, title in TABS:
            frm = ttk.Frame(self.nb)
            self.nb.add(frm, text=title.split(' (')[0] if key != 'balance' else title)
            hp = ttk.PanedWindow(frm, orient='horizontal')
            hp.pack(fill='both', expand=True)
            panel = SeriesPanel(hp, key, self)
            right = ttk.PanedWindow(hp, orient='vertical')
            hp.add(panel.frame, weight=0)
            hp.add(right, weight=1)
            f_fig = ttk.Frame(right)
            fig = Figure(figsize=(14, 6.5 if key == 'balance' else 5), layout='constrained')
            canvas = FigureCanvasTkAgg(fig, master=f_fig)
            NavigationToolbar2Tk(canvas, f_fig).update()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            f_tab = ttk.PanedWindow(right, orient='horizontal')
            trees = {}
            for tag in ('A', 'B'):
                tf = ttk.Frame(f_tab)
                lbl = ttk.Label(tf, text='', font=('Malgun Gothic', 9, 'bold'))
                lbl.pack(anchor='w')
                tree = ttk.Treeview(tf, show='headings')
                ys = ttk.Scrollbar(tf, orient='vertical', command=tree.yview)
                xs = ttk.Scrollbar(tf, orient='horizontal', command=tree.xview)
                tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
                ys.pack(side='right', fill='y')
                xs.pack(side='bottom', fill='x')
                tree.pack(fill='both', expand=True)
                trees[tag] = (tf, lbl, tree)
            f_tab.add(trees['A'][0], weight=1)
            right.add(f_fig, weight=3)
            right.add(f_tab, weight=1)
            self.tabs[key] = dict(fig=fig, canvas=canvas, panel=panel, tab_pw=f_tab, trees=trees, b_shown=False)
        self.nb.bind('<<NotebookTabChanged>>', lambda e: self._draw_current())

        self._refresh_files()
        root.protocol('WM_DELETE_WINDOW', self._close)

    # 설정
    def _load_cfg(self):
        try:
            with open(SETTINGS, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    def _close(self):
        try:
            tabs = {k: {kk: v for kk, v in st.items() if not kk.startswith('_')} for k, st in self.state.items()}
            with open(SETTINGS, 'w', encoding='utf-8') as f:
                json.dump({'folder': self.var_folder.get(), 'recursive': self.var_rec.get(),
                           'period': self.var_period.get(), 'sharey': self.var_sharey.get(),
                           'tabs': tabs}, f, ensure_ascii=False, indent=1)
        except Exception:
            pass
        self.root.destroy()

    # 파일
    def _pick_folder(self):
        from tkinter import filedialog
        d = filedialog.askdirectory(initialdir=self.var_folder.get())
        if d:
            self.var_folder.set(os.path.normpath(d))
            self._refresh_files()

    def _refresh_files(self):
        folder = self.var_folder.get()
        if not os.path.isdir(folder):
            self.var_status.set(f'폴더가 없습니다: {folder}')
            return
        self.files = core.list_result_files(folder, self.var_rec.get())
        names = [os.path.relpath(p, folder) for p in self.files]
        self.cb_file['A']['values'] = names
        self.cb_file['B']['values'] = [NO_FILE] + names
        if not self.var_file['B'].get():
            self.var_file['B'].set(NO_FILE)
        self.var_status.set(f'{len(self.files)}개 엑셀 파일 (최근 수정 순). 파일 A를 선택하면 불러옵니다. '
                            f'파일 B를 고르면 좌/우 비교 화면이 됩니다.')

    def _path_of(self, tag):
        i = self.cb_file[tag].current()
        if tag == 'B':
            i -= 1
        return self.files[i] if 0 <= i < len(self.files) else None

    def _load(self, tag, no_cache=False):
        path = self._path_of(tag)
        if path is None:
            if tag == 'B':
                self.objs['B'] = None
                self._apply()
            return
        self.var_status.set(f'[{tag}] 불러오는 중… {os.path.basename(path)} (처음 여는 파일은 20초 내외, 이후 캐시)')

        def work():
            try:
                obj = core.load_result(path, use_cache=not no_cache,
                                       progress=lambda m: self.q.put(('msg', tag, f'[{tag}] {m}')))
                self.q.put(('done', tag, obj))
            except Exception as ex:
                self.q.put(('err', tag, f'[{tag}] {type(ex).__name__}: {ex}'))
        threading.Thread(target=work, daemon=True).start()
        self.root.after(200, self._poll)

    def _poll(self):
        try:
            while True:
                kind, tag, val = self.q.get_nowait()
                if kind == 'msg':
                    self.var_status.set(val)
                elif kind == 'err':
                    self.var_status.set('오류: ' + val)
                    return
                elif kind == 'done':
                    first = self.objs['A'] is None and self.objs['B'] is None
                    self.objs[tag] = val
                    if first or not self.var_start.get():
                        hours = PERIODS[self.var_period.get()]
                        start = core.auto_start(val, AUTO_MODES[self.var_auto.get()], hours)
                        self.var_start.set(start.strftime('%Y-%m-%d %H:%M'))
                    self._apply()
                    return
        except queue.Empty:
            pass
        self.root.after(200, self._poll)

    def _swap(self):
        a, b = self.var_file['A'].get(), self.var_file['B'].get()
        if not a or b in ('', NO_FILE):
            return
        self.var_file['A'].set(b)
        self.var_file['B'].set(a)
        self.objs['A'], self.objs['B'] = self.objs['B'], self.objs['A']
        self._apply()

    def _loaded(self):
        return [(t, self.objs[t]) for t in ('A', 'B') if self.objs[t] is not None]

    # 구간
    def _ref(self):
        return self.objs['A'] or self.objs['B']

    def _apply(self):
        ref = self._ref()
        if ref is None:
            return
        try:
            start = pd.Timestamp(self.var_start.get())
        except Exception:
            self.var_status.set('분석 시작 형식: YYYY-MM-DD HH:MM')
            return
        hours = PERIODS[self.var_period.get()]
        self.sls = {t: core.window(o, start, hours) for t, o in self._loaded()}
        sl = core.window(ref, start, hours)
        idx = ref['index'][sl]
        self.var_start.set(idx[0].strftime('%Y-%m-%d %H:%M'))
        self.var_week.set(int(sl.start // 168) + 1)
        objs = [o for _, o in self._loaded()]
        for key, _ in TABS:
            merge_state(key, self.state, objs)
            self.tabs[key]['panel'].refresh()
        parts = []
        for t, o in self._loaded():
            res = o['balance']['residual'].iloc[self.sls[t]].abs().max()
            missing = [x.split(' ')[0] for k, x in core.SECTIONS if k not in o['sections']]
            parts.append(f'[{t}] {os.path.basename(o["path"])} (회랑: {o["sections"].get("_corridor_src", "-")}, '
                         f'수급오차 {res:.0e} MW' + (f', 원본 없음: {",".join(missing)}' if missing else '') + ')')
        self.var_status.set(f'{idx[0]:%Y-%m-%d %H:%M} ~ {idx[-1]:%Y-%m-%d %H:%M} ({len(idx)}h) | ' + ' | '.join(parts))
        self.redraw_all()

    def _apply_week(self):
        ref = self._ref()
        if ref is None:
            return
        idx = ref['index']
        loc = min(168 * (max(int(self.var_week.get()), 1) - 1), len(idx) - 1)
        self.var_start.set(idx[loc].strftime('%Y-%m-%d %H:%M'))
        self._apply()

    def _shift(self, k):
        ref = self._ref()
        if ref is None:
            return
        idx = ref['index']
        cur = int(idx.searchsorted(pd.Timestamp(self.var_start.get())))
        loc = min(max(cur + k * PERIODS[self.var_period.get()], 0), len(idx) - 1)
        self.var_start.set(idx[loc].strftime('%Y-%m-%d %H:%M'))
        self._apply()

    def _auto(self):
        ref = self._ref()
        if ref is None:
            return
        hours = PERIODS[self.var_period.get()]
        self.var_start.set(core.auto_start(ref, AUTO_MODES[self.var_auto.get()], hours).strftime('%Y-%m-%d %H:%M'))
        self._apply()

    # 그리기
    def redraw_all(self):
        self.dirty = set(self.tabs)
        self._draw_current()

    def redraw(self, key):
        self.dirty.add(key)
        if self._redraw_job:
            self.root.after_cancel(self._redraw_job)
        self._redraw_job = self.root.after(120, self._draw_current)

    def _current_key(self):
        return TABS[self.nb.index(self.nb.select())][0]

    def _panes(self):
        loaded = self._loaded()
        compare = len(loaded) > 1
        return [(pane_title(o, t if compare else None), o, self.sls[t]) for t, o in loaded if t in self.sls]

    def _draw_current(self):
        self._redraw_job = None
        panes = self._panes()
        if not panes:
            return
        key = self._current_key()
        if key not in self.dirty:
            return
        tab = self.tabs[key]
        st = self.state[key]
        plot_tab(tab['fig'], key, panes, st, sharey=self.var_sharey.get())
        tab['canvas'].draw_idle()
        compare = len(panes) > 1
        if compare and not tab['b_shown']:
            tab['tab_pw'].add(tab['trees']['B'][0], weight=1)
            tab['b_shown'] = True
        elif not compare and tab['b_shown']:
            tab['tab_pw'].forget(tab['trees']['B'][0])
            tab['b_shown'] = False
        loaded = [(t, o) for t, o in self._loaded() if t in self.sls]
        for t, o in loaded:
            _, lbl, tree = tab['trees'][t]
            lbl.configure(text=pane_title(o, t if compare else None))
            self._fill_tree(tree, tab_table(o, key, st, self.sls[t]))
        if not compare:
            other = 'B' if loaded[0][0] == 'A' else 'A'
            self.tabs[key]['trees'][other][2].delete(*self.tabs[key]['trees'][other][2].get_children())
            if loaded[0][0] == 'B':
                self._fill_tree(tab['trees']['A'][2], tab_table(loaded[0][1], key, st, self.sls['B']))
        self.dirty.discard(key)

    def _fill_tree(self, tree, df):
        tree.delete(*tree.get_children())
        cols = ['Timestamp'] + [str(c) for c in df.columns]
        tree['columns'] = cols
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=125 if c == 'Timestamp' else max(80, min(160, 8 * len(c))), anchor='e', stretch=False)
        for ts, vals in zip(df.index, df.itertuples(index=False)):
            tree.insert('', 'end', values=[ts.strftime('%Y-%m-%d %H:%M')] + [f'{v:,.1f}' for v in vals])

    # 저장
    def _default_name(self, obj, sl, ext):
        idx = obj['index'][sl]
        stem = os.path.splitext(os.path.basename(obj['path']))[0]
        return f'AW_{stem}_{idx[0]:%m%d%H}_{len(idx)}h{ext}'

    def _save_png(self):
        from tkinter import filedialog
        loaded = [(t, o) for t, o in self._loaded() if t in self.sls]
        if not loaded:
            return
        key = self._current_key()
        t, o = loaded[0]
        name = self._default_name(o, self.sls[t], f'_{key}{"_비교" if len(loaded) > 1 else ""}.png')
        p = filedialog.asksaveasfilename(initialdir=os.path.dirname(o['path']), defaultextension='.png',
                                         initialfile=name, filetypes=[('PNG', '*.png')])
        if p:
            self.tabs[key]['fig'].savefig(p, dpi=150)
            self.var_status.set(f'그림 저장: {p}')

    def _export(self):
        from tkinter import filedialog, messagebox
        loaded = [(t, o) for t, o in self._loaded() if t in self.sls]
        if not loaded:
            return
        t0, o0 = loaded[0]
        if len(loaded) == 1:
            p = filedialog.asksaveasfilename(initialdir=os.path.dirname(o0['path']), defaultextension='.xlsx',
                                             initialfile=self._default_name(o0, self.sls[t0], '.xlsx'),
                                             filetypes=[('Excel', '*.xlsx')])
            targets = [(o0, self.sls[t0], p)] if p else []
        else:
            d = filedialog.askdirectory(initialdir=os.path.dirname(o0['path']),
                                        title='A·B 두 파일을 저장할 폴더 선택')
            targets = [(o, self.sls[t], os.path.join(d, self._default_name(o, self.sls[t], f'_{t}.xlsx')))
                       for t, o in loaded] if d else []
        if not targets:
            return
        saved = []
        for o, sl, p in targets:
            try:
                saved.append(export_excel(o, sl, p, self.state['balance']))
            except PermissionError:
                messagebox.showerror('엑셀 내보내기', f'파일이 열려 있어 저장할 수 없습니다. 닫고 다시 시도하세요.\n{p}')
                return
        self.var_status.set('엑셀 저장: ' + ' | '.join(saved))
        if messagebox.askyesno('엑셀 내보내기', '저장했습니다.\n' + '\n'.join(saved) + '\n\n지금 열까요?'):
            for p in saved:
                os.startfile(p)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('folder', nargs='?', help='GUI 시작 폴더')
    ap.add_argument('--file', help='배치 모드: 결과 엑셀 경로')
    ap.add_argument('--compare', help='배치 모드: 비교할 두 번째 결과 엑셀')
    ap.add_argument('--start', help='분석 시작 (YYYY-MM-DD HH:MM)')
    ap.add_argument('--hours', type=int, default=168)
    ap.add_argument('--auto', choices=sorted(set(AUTO_MODES.values())))
    ap.add_argument('--legend', default=LEGEND_POS[0], choices=LEGEND_POS)
    ap.add_argument('--png')
    ap.add_argument('--all-png', action='store_true', help='7개 섹션 그림도 저장')
    ap.add_argument('--export')
    ap.add_argument('--no-cache', action='store_true')
    a = ap.parse_args()
    if a.file:
        run_batch(a)
        return
    import tkinter as tk
    root = tk.Tk()
    ViewerApp(root, os.path.abspath(a.folder) if a.folder else None)
    root.mainloop()


if __name__ == '__main__':
    sys.exit(main())
