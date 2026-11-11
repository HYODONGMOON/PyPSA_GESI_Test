"""
Analysis Window 뷰어 — 기존 결과 엑셀을 재분석 없이 시간별로 점검

사용법
------
  GUI   : python analysis_window_viewer.py [결과폴더]
  배치  : python analysis_window_viewer.py --file 결과.xlsx --start "2024-08-01 00:00" --hours 168
                 [--auto line|curtail|demand|net_low] [--png 그림.png] [--export AW_결과.xlsx]

폴더 안의 결과 파일(*.xlsx) 중 하나를 골라 지정 기간의
  · 전력수급 그림 (Figure_hour_electricity 형식: 공급 누적 + 수요선 / 유연 소비 누적)
  · Analysis_Window 7개 섹션 (송전 회랑, 발전원, 링크, 저장량, 저장 출입, Curtailment, 수요)
을 보여주고, 현재 구간을 엑셀(차트 포함)로 내보낸다.
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


# ── 그림 ─────────────────────────────────────────────────────────────
def _mpl_setup():
    import matplotlib
    matplotlib.rcParams['font.family'] = 'Malgun Gothic'
    matplotlib.rcParams['axes.unicode_minus'] = False


def _fmt_time_axis(ax, idx):
    import matplotlib.dates as mdates
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


def _nonzero(df, tol=1e-3):
    return df.loc[:, df.abs().max() > tol]


def plot_balance(fig, obj, sl):
    """Figure_hour_electricity: 위=공급 누적+수요선, 아래=유연 소비 누적."""
    fig.clear()
    b = obj['balance']
    sup = _nonzero(b['supply'].iloc[sl]).clip(lower=0)
    con = _nonzero(b['consume'].iloc[sl]).clip(lower=0)
    dem = b['demand'].iloc[sl]
    idx = sup.index
    lab_s = {k: (l, c) for k, l, c in core.SUPPLY_DEF}
    lab_c = {k: (l, c) for k, l, c in core.CONSUME_DEF}

    ax1, ax2 = fig.subplots(2, 1, sharex=True, gridspec_kw={'height_ratios': [3, 2]})
    ax1.stackplot(idx, *[sup[k].values for k in sup],
                  labels=[lab_s.get(k, (k, None))[0] for k in sup],
                  colors=[lab_s.get(k, (k, '#999999'))[1] for k in sup], alpha=0.9, linewidth=0)
    ax1.plot(idx, dem.values, color='black', lw=1.6, label=core.DEMAND_LABEL)
    flex = con.drop(columns=['Curtail'], errors='ignore').sum(axis=1)
    ax1.plot(idx, (dem + flex).values, color='black', lw=1.0, ls='--', label='전력수요+유연소비')
    ax1.set_ylabel('공급 (MW)')
    ax1.legend(loc='upper left', bbox_to_anchor=(1.005, 1.0), fontsize=8, frameon=False)
    ax1.set_title('시간별 전력수급 (국가, 전력모선 기준)', fontsize=11, loc='left')

    ax2.stackplot(idx, *[con[k].values for k in con],
                  labels=[lab_c.get(k, (k, None))[0] for k in con],
                  colors=[lab_c.get(k, (k, '#999999'))[1] for k in con], alpha=0.9, linewidth=0)
    ax2.set_ylabel('유연 소비 (MW)')
    ax2.legend(loc='upper left', bbox_to_anchor=(1.005, 1.0), fontsize=8, frameon=False)
    for ax in (ax1, ax2):
        _fmt_time_axis(ax, idx)
        ax.yaxis.set_major_formatter(__import__('matplotlib').ticker.FuncFormatter(lambda v, _: f'{v:,.0f}'))
    fig.subplots_adjust(left=0.07, right=0.83, top=0.94, bottom=0.09, hspace=0.08)


def plot_section(fig, obj, key, sl):
    fig.clear()
    ax = fig.add_subplot(111)
    sec = obj['sections'].get(key)
    if sec is None:
        ax.text(0.5, 0.5, '원본 시트가 없어 표시할 수 없습니다', ha='center', va='center', transform=ax.transAxes)
        return
    df = sec.iloc[sl]
    idx = df.index
    if key == 'gen':
        keep = [c for c in obj['balance']['supply'].columns if c in df.columns]
        d = _nonzero(df[keep]).clip(lower=0)
        lab = {k: (l, c) for k, l, c in core.SUPPLY_DEF}
        ax.stackplot(idx, *[d[k].values for k in d], labels=[lab.get(k, (k,))[0] for k in d],
                     colors=[lab.get(k, (k, '#999999'))[1] for k in d], alpha=0.9, linewidth=0)
        ax.set_title('발전원별 출력 (전력모선 발전기, 누적)', fontsize=11, loc='left')
    elif key == 'curtail':
        pv = df[[c for c in df.columns if c.endswith('_PV')]].sum(axis=1)
        wt = df[[c for c in df.columns if c.endswith('_WT')]].sum(axis=1)
        ax.stackplot(idx, pv.values, wt.values, labels=['태양광', '풍력'],
                     colors=['#FFD966', '#70AD47'], alpha=0.9, linewidth=0)
        ax.set_title('재생에너지 Curtailment (국가, 누적) — 발전기별 값은 아래 표', fontsize=11, loc='left')
    else:
        import matplotlib
        d = _nonzero(df)
        cmap = matplotlib.colormaps['tab20']
        for i, c in enumerate(d):
            ax.plot(idx, d[c].values, lw=1.2, label=c, color=cmap(i % 20))
        if key == 'corridor':
            ax.axhline(0, color='black', lw=0.8)
            ax.set_title(f'지역간 회랑 조류 (+: 열 이름의 화살표 방향)  [{obj["sections"].get("_corridor_src", "")}]',
                         fontsize=11, loc='left')
    unit = 'MWh' if key in ('store_e', 'curtail') else 'MW'
    ax.set_ylabel(unit)
    ax.legend(loc='upper left', bbox_to_anchor=(1.005, 1.0), fontsize=8, frameon=False,
              ncol=2 if len(ax.get_legend_handles_labels()[0]) > 14 else 1)
    _fmt_time_axis(ax, idx)
    ax.yaxis.set_major_formatter(__import__('matplotlib').ticker.FuncFormatter(lambda v, _: f'{v:,.0f}'))
    fig.subplots_adjust(left=0.07, right=0.80, top=0.93, bottom=0.12)


# ── 엑셀 내보내기 ────────────────────────────────────────────────────
def export_excel(obj, sl, out_path):
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

    ch1 = _area(sup_cols, '시간별 전력수급 — 공급(누적)과 전력수요 (MW)')
    ln = LineChart()
    j = col_order.index(dem_col) + 1
    ln.add_data(Reference(ws, min_col=j, min_row=1, max_row=n + 1), titles_from_data=True)
    ln.series[0].graphicalProperties.line.solidFill = '000000'
    ln.series[0].graphicalProperties.line.width = 22000
    ln.series[0].smooth = False
    ch1 += ln
    ch2 = _area(con_cols, '유연 소비(누적) 및 Curtailment (MW)')
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
    obj = core.load_result(a.file, use_cache=not a.no_cache, progress=print)
    start = core.auto_start(obj, a.auto, a.hours) if a.auto else (a.start or obj['index'][0])
    sl = core.window(obj, start, a.hours)
    idx = obj['index'][sl]
    print(f'구간: {idx[0]} ~ {idx[-1]} ({len(idx)}h), 회랑 데이터: {obj["sections"].get("_corridor_src")}')
    print(f'수급 검증 |공급-소비| 최대 {obj["balance"]["residual"].iloc[sl].abs().max():.2e} MW')
    if a.png:
        _mpl_setup()
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        fig = plt.figure(figsize=(15, 8))
        plot_balance(fig, obj, sl)
        fig.savefig(a.png, dpi=130)
        print('그림 저장:', a.png)
        for key, _ in core.SECTIONS if a.all_png else []:
            fig = plt.figure(figsize=(15, 5.5))
            plot_section(fig, obj, key, sl)
            p = os.path.splitext(a.png)[0] + f'_{key}.png'
            fig.savefig(p, dpi=110)
            print('그림 저장:', p)
    if a.export:
        print('엑셀 저장:', export_excel(obj, sl, a.export))


# ── GUI ──────────────────────────────────────────────────────────────
class ViewerApp:
    def __init__(self, root, folder=None):
        import tkinter as tk
        from tkinter import ttk
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
        from matplotlib.figure import Figure
        _mpl_setup()
        self.tk, self.ttk = tk, ttk
        self.root = root
        self.obj = None
        self.sl = None
        self.files = []
        self.q = queue.Queue()
        self.dirty = set()
        cfg = self._load_cfg()

        root.title('Analysis Window 뷰어 — PyPSA 결과 시간별 점검')
        root.geometry('1500x950')

        top = ttk.Frame(root, padding=6)
        top.pack(fill='x')
        r1 = ttk.Frame(top)
        r1.pack(fill='x', pady=2)
        ttk.Label(r1, text='결과 폴더').pack(side='left')
        self.var_folder = tk.StringVar(value=folder or cfg.get('folder') or BASE_DIR)
        ttk.Entry(r1, textvariable=self.var_folder, width=110).pack(side='left', padx=4, fill='x', expand=True)
        ttk.Button(r1, text='폴더 선택', command=self._pick_folder).pack(side='left')
        self.var_rec = tk.BooleanVar(value=cfg.get('recursive', True))
        ttk.Checkbutton(r1, text='하위 폴더 포함', variable=self.var_rec, command=self._refresh_files).pack(side='left', padx=6)
        ttk.Button(r1, text='새로고침', command=self._refresh_files).pack(side='left')

        r2 = ttk.Frame(top)
        r2.pack(fill='x', pady=2)
        ttk.Label(r2, text='결과 파일').pack(side='left')
        self.var_file = tk.StringVar()
        self.cb_file = ttk.Combobox(r2, textvariable=self.var_file, width=120, state='readonly')
        self.cb_file.pack(side='left', padx=4, fill='x', expand=True)
        self.cb_file.bind('<<ComboboxSelected>>', lambda e: self._load())
        ttk.Button(r2, text='불러오기', command=self._load).pack(side='left')
        ttk.Button(r2, text='캐시 무시 재읽기', command=lambda: self._load(no_cache=True)).pack(side='left', padx=4)

        r3 = ttk.Frame(top)
        r3.pack(fill='x', pady=2)
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
        ttk.Label(r3, text='자동 구간').pack(side='left', padx=(12, 0))
        self.var_auto = tk.StringVar(value='선로 포화 최대')
        ttk.Combobox(r3, textvariable=self.var_auto, values=list(AUTO_MODES), width=20,
                     state='readonly').pack(side='left', padx=4)
        ttk.Button(r3, text='찾기', command=self._auto).pack(side='left')
        ttk.Button(r3, text='엑셀 내보내기', command=self._export).pack(side='right')
        ttk.Button(r3, text='그림 저장(PNG)', command=self._save_png).pack(side='right', padx=4)

        self.var_status = tk.StringVar(value='결과 파일을 선택하세요.')
        ttk.Label(root, textvariable=self.var_status, anchor='w', relief='sunken', padding=3).pack(side='bottom', fill='x')

        self.nb = ttk.Notebook(root)
        self.nb.pack(fill='both', expand=True)
        self.tabs = {}
        for key, title in [('balance', '전력수급 (Figure_hour_electricity)')] + core.SECTIONS:
            frm = ttk.Frame(self.nb)
            self.nb.add(frm, text=title.split(' (')[0] if key != 'balance' else title)
            pw = ttk.PanedWindow(frm, orient='vertical')
            pw.pack(fill='both', expand=True)
            f_fig = ttk.Frame(pw)
            fig = Figure(figsize=(14, 6.5 if key == 'balance' else 5))
            canvas = FigureCanvasTkAgg(fig, master=f_fig)
            NavigationToolbar2Tk(canvas, f_fig).update()
            canvas.get_tk_widget().pack(fill='both', expand=True)
            f_tab = ttk.Frame(pw)
            tree = ttk.Treeview(f_tab, show='headings')
            ys = ttk.Scrollbar(f_tab, orient='vertical', command=tree.yview)
            xs = ttk.Scrollbar(f_tab, orient='horizontal', command=tree.xview)
            tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
            ys.pack(side='right', fill='y')
            xs.pack(side='bottom', fill='x')
            tree.pack(fill='both', expand=True)
            pw.add(f_fig, weight=3)
            pw.add(f_tab, weight=1)
            self.tabs[key] = (fig, canvas, tree)
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
            with open(SETTINGS, 'w', encoding='utf-8') as f:
                json.dump({'folder': self.var_folder.get(), 'recursive': self.var_rec.get(),
                           'period': self.var_period.get()}, f, ensure_ascii=False, indent=1)
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
        self.cb_file['values'] = [os.path.relpath(p, folder) for p in self.files]
        self.var_status.set(f'{len(self.files)}개 엑셀 파일 (최근 수정 순). 파일을 선택하면 불러옵니다.')

    def _load(self, no_cache=False):
        i = self.cb_file.current()
        if i < 0:
            return
        path = self.files[i]
        self.var_status.set(f'불러오는 중… {os.path.basename(path)} (처음 여는 파일은 20초 내외, 이후 캐시)')

        def work():
            try:
                obj = core.load_result(path, use_cache=not no_cache,
                                       progress=lambda m: self.q.put(('msg', m)))
                self.q.put(('done', obj))
            except Exception as ex:
                self.q.put(('err', f'{type(ex).__name__}: {ex}'))
        threading.Thread(target=work, daemon=True).start()
        self.root.after(200, self._poll)

    def _poll(self):
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == 'msg':
                    self.var_status.set(val)
                elif kind == 'err':
                    self.var_status.set('오류: ' + val)
                    return
                elif kind == 'done':
                    self.obj = val
                    hours = PERIODS[self.var_period.get()]
                    start = core.auto_start(val, AUTO_MODES[self.var_auto.get()], hours)
                    self.var_start.set(start.strftime('%Y-%m-%d %H:%M'))
                    self._apply()
                    return
        except queue.Empty:
            pass
        self.root.after(200, self._poll)

    # 구간
    def _apply(self):
        if self.obj is None:
            return
        try:
            start = pd.Timestamp(self.var_start.get())
        except Exception:
            self.var_status.set('분석 시작 형식: YYYY-MM-DD HH:MM')
            return
        self.sl = core.window(self.obj, start, PERIODS[self.var_period.get()])
        idx = self.obj['index'][self.sl]
        self.var_start.set(idx[0].strftime('%Y-%m-%d %H:%M'))
        self.var_week.set(int(self.sl.start // 168) + 1)
        res = self.obj['balance']['residual'].iloc[self.sl].abs().max()
        missing = [t.split(' ')[0] for k, t in core.SECTIONS if k not in self.obj['sections']]
        self.var_status.set(
            f'{os.path.basename(self.obj["path"])} | {idx[0]:%Y-%m-%d %H:%M} ~ {idx[-1]:%Y-%m-%d %H:%M} ({len(idx)}h) | '
            f'회랑: {self.obj["sections"].get("_corridor_src", "-")} | 수급검증 최대오차 {res:.1e} MW'
            + (f' | 원본 없음: {", ".join(missing)}' if missing else ''))
        self.dirty = set(self.tabs)
        self._draw_current()

    def _apply_week(self):
        if self.obj is None:
            return
        idx = self.obj['index']
        loc = min(168 * (max(int(self.var_week.get()), 1) - 1), len(idx) - 1)
        self.var_start.set(idx[loc].strftime('%Y-%m-%d %H:%M'))
        self._apply()

    def _shift(self, k):
        if self.obj is None or self.sl is None:
            return
        idx = self.obj['index']
        loc = min(max(self.sl.start + k * PERIODS[self.var_period.get()], 0), len(idx) - 1)
        self.var_start.set(idx[loc].strftime('%Y-%m-%d %H:%M'))
        self._apply()

    def _auto(self):
        if self.obj is None:
            return
        hours = PERIODS[self.var_period.get()]
        self.var_start.set(core.auto_start(self.obj, AUTO_MODES[self.var_auto.get()], hours).strftime('%Y-%m-%d %H:%M'))
        self._apply()

    # 그리기
    def _current_key(self):
        return list(self.tabs)[self.nb.index(self.nb.select())]

    def _draw_current(self):
        if self.obj is None or self.sl is None:
            return
        key = self._current_key()
        if key not in self.dirty:
            return
        fig, canvas, tree = self.tabs[key]
        if key == 'balance':
            plot_balance(fig, self.obj, self.sl)
            df = core.balance_table(self.obj, self.sl)
        else:
            plot_section(fig, self.obj, key, self.sl)
            sec = self.obj['sections'].get(key)
            df = sec.iloc[self.sl] if sec is not None else pd.DataFrame()
        canvas.draw_idle()
        self._fill_tree(tree, df)
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
    def _default_name(self, ext):
        idx = self.obj['index'][self.sl]
        stem = os.path.splitext(os.path.basename(self.obj['path']))[0]
        return f'AW_{stem}_{idx[0]:%m%d%H}_{len(idx)}h{ext}'

    def _save_png(self):
        from tkinter import filedialog
        if self.obj is None:
            return
        key = self._current_key()
        p = filedialog.asksaveasfilename(initialdir=os.path.dirname(self.obj['path']), defaultextension='.png',
                                         initialfile=self._default_name(f'_{key}.png'), filetypes=[('PNG', '*.png')])
        if p:
            self.tabs[key][0].savefig(p, dpi=150)
            self.var_status.set(f'그림 저장: {p}')

    def _export(self):
        from tkinter import filedialog, messagebox
        if self.obj is None:
            return
        p = filedialog.asksaveasfilename(initialdir=os.path.dirname(self.obj['path']), defaultextension='.xlsx',
                                         initialfile=self._default_name('.xlsx'), filetypes=[('Excel', '*.xlsx')])
        if not p:
            return
        try:
            export_excel(self.obj, self.sl, p)
            self.var_status.set(f'엑셀 저장: {p}')
            if messagebox.askyesno('엑셀 내보내기', f'저장했습니다.\n{p}\n\n지금 열까요?'):
                os.startfile(p)
        except PermissionError:
            messagebox.showerror('엑셀 내보내기', '파일이 열려 있어 저장할 수 없습니다. 닫고 다시 시도하세요.')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('folder', nargs='?', help='GUI 시작 폴더')
    ap.add_argument('--file', help='배치 모드: 결과 엑셀 경로')
    ap.add_argument('--start', help='분석 시작 (YYYY-MM-DD HH:MM)')
    ap.add_argument('--hours', type=int, default=168)
    ap.add_argument('--auto', choices=sorted(set(AUTO_MODES.values())))
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
    ViewerApp(root, a.folder)
    root.mainloop()


if __name__ == '__main__':
    sys.exit(main())
