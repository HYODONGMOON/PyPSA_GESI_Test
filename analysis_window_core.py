"""
Analysis Window 데이터 처리 (결과 엑셀 → 시간별 분석 데이터)

PyPSA_GUI.save_results 가 만든 결과 파일(optimization_result_*.xlsx)에서
Analysis_Window 7개 섹션과 전력수급(Figure_hour_electricity) 데이터를 만든다.
파생 시트(Corridor_Line_Flow 등)가 없으면 원본 시트(Line_Flow + Line_Info)에서 계산한다.
"""
import os
import pickle
import re

import numpy as np
import pandas as pd

CACHE_VERSION = 3
CACHE_DIRNAME = '_aw_cache'
REGIONS = {'BSN', 'CBD', 'CND', 'GBD', 'GGD', 'GND', 'GWD', 'ICN', 'JBD', 'JJD', 'JND'}

RAW_SHEETS = [
    'Corridor_Line_Flow', 'Line_Flow', 'Line_Info',
    'National_Gen_Output', 'Generator_Info',
    'National_Link_Flow', 'Link_Flow', 'Link_Info',
    'National_Storage_Energy', 'National_Storage_Power', 'Storage_Power', 'Storage_Info',
    'RE_Curtailment_Timeseries', 'National_Loads', 'Hourly_Loads',
]
INFO_SHEETS = {'Line_Info', 'Generator_Info', 'Link_Info', 'Storage_Info'}

# (key, 섹션 제목) — PyPSA_GUI._build_analysis_window_sheet 의 SECTIONS_DEF 순서
SECTIONS = [
    ('corridor', '1. 송전망 조류 현황 (지역간 연결선 통합, MW)'),
    ('gen', '2. 발전원별 출력 현황 (MW)'),
    ('link', '3. 유연성자원 / 섹터커플링 운영현황 (링크 조류 MW)'),
    ('store_e', '4. 저장설비 현황(저장량) (MWh)'),
    ('store_p', '5. 저장설비 현황(출입) (MW)'),
    ('curtail', '6. 재생에너지 Curtailment (MWh)'),
    ('loads', '7. 국가 전체 수요 현황 (전력/열/EV/Fab/DC, MW)'),
]

# 전력수급 그림 계열: (key, 표시명, 색)
SUPPLY_DEF = [
    ('Nuclear', '원자력', '#7B3F9E'),
    ('Coal', '석탄', '#595959'),
    ('LNG', 'LNG', '#F4A259'),
    ('CHP', 'CHP', '#B5651D'),
    ('Hydro', '수력', '#2E75B6'),
    ('PV', '태양광', '#FFD966'),
    ('WT', '풍력', '#70AD47'),
    ('Fuelcell', '연료전지', '#00B0B0'),
    ('ESS_Disch', 'ESS 방전', '#E06666'),
    ('Pumped_out', '양수 발전', '#1F4E79'),
    ('EV_V2G', 'V2G', '#FF99CC'),
    ('Heat_Disch', '열저장 방전(전력)', '#C9A0DC'),
    ('DR', 'DR', '#BFBFBF'),
    ('slack', '미공급(slack)', '#000000'),
]
CONSUME_DEF = [
    ('Curtail', 'Curtailment', '#D9D9D9'),
    ('HP', '히트펌프(P2H)', '#C00000'),
    ('Electrolyser', '수전해', '#00B050'),
    ('EV_Charger', 'EV 충전', '#FF66CC'),
    ('ESS_Ch', 'ESS 충전', '#E06666'),
    ('Pumped_in', '양수 충전', '#1F4E79'),
    ('EL_to_DC_Fab_RE', 'Fab·DC 공급', '#8064A2'),
    ('Heat_Ch', '열저장 충전(전력)', '#C9A0DC'),
    ('HVDC', 'HVDC 손실', '#7F7F7F'),
]
DEMAND_KEY = 'el_demand'
DEMAND_LABEL = '전력수요(Fab·DC 포함)'


def _region(bus):
    return str(bus).split('_')[0]


def _bus_type(bus):
    s = str(bus)
    head = s.split('_', 1)
    return head[1] if len(head) == 2 and head[0] in REGIONS else s


def _comp_type(name):
    head = str(name).split('_', 1)
    return head[1] if len(head) == 2 and head[0] in REGIONS else str(name)


def _is_el(bus):
    return _bus_type(bus) == 'EL'


# ── 엑셀 읽기 ─────────────────────────────────────────────────────────
def _read_sheets(path, wanted, progress=None):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    try:
        present = [s for s in wanted if s in wb.sheetnames]
        for i, s in enumerate(present):
            if progress:
                progress(f'시트 읽는 중 ({i + 1}/{len(present)}): {s}')
            rows = wb[s].iter_rows(values_only=True)
            header = next(rows, None)
            if header is None:
                continue
            data = [r for r in rows if r and r[0] is not None]
            cols = [str(c).strip() if c is not None else f'col{j}' for j, c in enumerate(header)]
            df = pd.DataFrame(data, columns=cols)
            df = df.set_index(cols[0])
            if s in INFO_SHEETS:
                df.index = df.index.map(lambda v: str(v).strip())
            else:
                df.index = pd.to_datetime(df.index)
                df = df.apply(pd.to_numeric, errors='coerce').fillna(0.0)
            out[s] = df
        out['_sheetnames'] = list(wb.sheetnames)
    finally:
        wb.close()
    return out


COAL_PLANT_KEYWORDS = ['당진', '영흥', '하동', '삼척', '태안', '보령', '삼천포', '호남', '동해',
                       '강릉안인', '북평', '신서천', '고성', '여수', '삼척그린파워']


def _open_network(path, progress=None):
    """netCDF 는 비ASCII 경로를 열지 못하므로 필요하면 임시 폴더로 복사해서 연다."""
    import shutil
    import tempfile
    import pypsa
    if progress:
        progress('PyPSA 네트워크(.nc) 여는 중')
    src = path
    tmp = None
    if not path.isascii():
        tmp = os.path.join(tempfile.gettempdir(), f'aw_nc_{abs(hash(path))}.nc')
        shutil.copyfile(path, tmp)
        src = tmp
    try:
        return pypsa.Network(src)
    finally:
        if tmp and os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def _raw_from_network(path, progress=None):
    """PyPSA .nc → save_results 가 만드는 시트와 같은 구조의 raw dict."""
    n = _open_network(path, progress)
    if progress:
        progress('네트워크에서 시간별 데이터 구성 중')

    def clean(df):
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        return df.fillna(0.0)

    def info(df, cols):
        df = df[[c for c in cols if c in df.columns]].copy()
        df.index = [str(i).strip() for i in df.index]
        return df

    raw = {}
    gtp = n.generators_t.p
    if gtp is None or gtp.empty:
        return {'_sheetnames': []}
    coal = [c for c in gtp.columns
            if 'coal' in str(n.generators.at[c, 'carrier']).lower() or any(k in c for k in COAL_PLANT_KEYWORDS)]
    agg = gtp.drop(columns=coal)
    for col in coal:
        k = f"{col.split('_')[0] if '_' in col else 'ETC'}_Coal"
        agg[k] = agg[k] + gtp[col] if k in agg else gtp[col]
    raw['National_Gen_Output'] = _national(clean(agg))
    raw['Generator_Info'] = info(n.generators, ['bus', 'p_nom', 'p_nom_opt'])

    if not n.lines_t.p0.empty:
        raw['Line_Flow'] = clean(n.lines_t.p0)
        raw['Line_Info'] = info(n.lines, ['bus0', 'bus1', 's_nom', 's_nom_opt'])
    if not n.links_t.p0.empty:
        raw['Link_Flow'] = clean(n.links_t.p0)
        raw['National_Link_Flow'] = _national(raw['Link_Flow'])
        raw['Link_Info'] = info(n.links, ['bus0', 'bus1', 'p_nom', 'efficiency', 'p_nom_opt'])
    if not n.stores_t.p.empty:
        raw['Storage_Power'] = clean(n.stores_t.p)
        raw['National_Storage_Power'] = _national(raw['Storage_Power'])
        raw['Storage_Info'] = info(n.stores, ['bus', 'carrier', 'e_nom', 'e_nom_opt'])
    if not n.stores_t.e.empty:
        raw['National_Storage_Energy'] = _national(clean(n.stores_t.e))

    curt = {}
    pmax = n.generators_t.p_max_pu
    for g in n.generators.index:
        if not any(k in g.lower() for k in ['pv', 'solar', 'wind', 'wt']):
            continue
        if g not in pmax.columns or g not in gtp.columns:
            continue
        opt = n.generators.at[g, 'p_nom_opt'] if 'p_nom_opt' in n.generators else float('nan')
        ext = bool(n.generators.at[g, 'p_nom_extendable'])
        p_nom = float(opt) if ext and pd.notna(opt) and float(opt) > 0 else float(n.generators.at[g, 'p_nom'])
        c = (pmax[g] * p_nom - gtp[g]).clip(lower=0)
        if c.sum() > 0.1:
            curt[g.strip()] = c
    if curt:
        raw['RE_Curtailment_Timeseries'] = pd.DataFrame(curt, index=gtp.index)

    ltp = n.loads_t.p if not n.loads_t.p.empty else n.loads_t.p_set
    if ltp is not None and not ltp.empty:
        hl = clean(ltp)
        raw['Hourly_Loads'] = hl
        cols = list(hl.columns)
        groups = {
            '전력수요_MW': [c for c in cols if '_Demand_EL' in c or (c.endswith('_EL') and '_Demand_' not in c)],
            '열수요_MW': [c for c in cols if '_Demand_H' in c and '_H2' not in c and '_EL' not in c and '_EV' not in c],
            'EV수요_MW': [c for c in cols if '_Demand_EV' in c or (c.endswith('_EV') and '_Demand_' not in c)],
            'Fab수요_MW': [c for c in cols if '_Demand_Fab' in c],
            'DC수요_MW': [c for c in cols if '_Demand_DC' in c],
        }
        raw['National_Loads'] = pd.DataFrame({k: hl[v].sum(axis=1) for k, v in groups.items() if v}, index=hl.index)
    raw['_sheetnames'] = [k for k in raw]
    return raw


def _national(df):
    agg = {}
    for c in df.columns:
        k = _comp_type(c)
        agg[k] = agg[k] + df[c] if k in agg else df[c].copy()
    return pd.DataFrame(agg, index=df.index)


# ── 파생 데이터 ───────────────────────────────────────────────────────
def build_corridor(line_flow, line_info):
    """지역간 회랑 조류. 열 이름 'A→B' 에서 양(+)은 A→B 방향."""
    corr, orient = {}, {}
    for line in line_flow.columns:
        if line not in line_info.index:
            continue
        r0, r1 = _region(line_info.at[line, 'bus0']), _region(line_info.at[line, 'bus1'])
        if r0 == r1:
            continue
        key = frozenset((r0, r1))
        if key not in orient:
            orient[key] = (r0, r1)
        sign = 1.0 if orient[key] == (r0, r1) else -1.0
        p = line_flow[line] * sign
        corr[key] = corr[key] + p if key in corr else p.copy()
    return pd.DataFrame({f'{orient[k][0]}→{orient[k][1]}': v for k, v in corr.items()},
                        index=line_flow.index)


def build_el_balance(raw):
    """국가 전력모선(EL) 수급: 공급(+)·소비(+) 계열과 전력수요."""
    idx = raw['National_Gen_Output'].index
    sup, con = {}, {}

    def add(d, k, s):
        d[k] = d[k] + s if k in d else s.copy()

    gi = raw.get('Generator_Info')
    non_el = set()
    if gi is not None:
        for g, b in gi['bus'].items():
            if not _is_el(b):
                non_el.add(_comp_type(g))
    for c in raw['National_Gen_Output'].columns:
        if c in non_el:
            continue
        add(sup, c, raw['National_Gen_Output'][c])

    lf, li = raw.get('Link_Flow'), raw.get('Link_Info')
    if lf is not None and li is not None:
        for L in lf.columns:
            if L not in li.index:
                continue
            b0, b1 = li.at[L, 'bus0'], li.at[L, 'bus1']
            eff = float(li.at[L, 'efficiency']) if pd.notna(li.at[L, 'efficiency']) else 1.0
            p0 = lf[L]
            t = _comp_type(L)
            if _is_el(b0) and _is_el(b1):
                add(con, t, p0 * (1.0 - eff))
            elif _is_el(b1):
                add(sup, t, p0 * eff)
            elif _is_el(b0):
                add(con, t, p0)

    sp, si = raw.get('Storage_Power'), raw.get('Storage_Info')
    if sp is not None and si is not None:
        for s in sp.columns:
            if s in si.index and _is_el(si.at[s, 'bus']):
                t = _comp_type(s)
                add(sup, f'{t}_out', sp[s].clip(lower=0))
                add(con, f'{t}_in', (-sp[s]).clip(lower=0))

    if 'RE_Curtailment_Timeseries' in raw:
        con['Curtail'] = raw['RE_Curtailment_Timeseries'].sum(axis=1)

    demand_parts = el_bus_demand(raw, idx)
    demand = demand_parts.sum(axis=1)

    sup_df = pd.DataFrame(sup, index=idx)
    con_df = pd.DataFrame(con, index=idx)
    order_s = [k for k, _, _ in SUPPLY_DEF if k in sup_df] + [k for k in sup_df if k not in dict((a, 1) for a, _, _ in SUPPLY_DEF)]
    order_c = [k for k, _, _ in CONSUME_DEF if k in con_df] + [k for k in con_df if k not in dict((a, 1) for a, _, _ in CONSUME_DEF)]
    sup_df, con_df = sup_df[order_s], con_df[order_c]
    resid = sup_df.sum(axis=1) - demand - con_df.drop(columns=['Curtail'], errors='ignore').sum(axis=1)
    return {'supply': sup_df, 'consume': con_df, 'demand': demand.rename(DEMAND_KEY),
            'demand_parts': demand_parts, 'residual': resid}


def el_bus_demand(raw, idx):
    """전력모선(EL)에 직접 연결된 부하: 일반 전력수요 + 계통 공급 DC/Fab.
    '_RE' 로 끝나는 DC/Fab 부하는 DC_Fab_RE 모선(EL_to_DC_Fab_RE 링크로 공급)이라 제외."""
    hl = raw.get('Hourly_Loads')
    parts = {}
    if hl is not None:
        for c in hl.columns:
            t = _comp_type(c.strip())
            if t.endswith('_RE'):
                continue
            label = {'Demand_EL': '일반 전력수요', 'Demand_DC': 'DC 수요(계통)',
                     'Demand_Fab': 'Fab 수요(계통)'}.get(t)
            if label:
                parts[label] = parts[label] + hl[c] if label in parts else hl[c].copy()
    else:
        nl = raw.get('National_Loads')
        if nl is not None and '전력수요_MW' in nl:
            parts['일반 전력수요'] = nl['전력수요_MW']
    if not parts:
        parts['일반 전력수요'] = pd.Series(0.0, index=idx)
    return pd.DataFrame(parts, index=idx)


def build_sections(raw):
    sec = {}
    if 'Corridor_Line_Flow' in raw:
        sec['corridor'] = raw['Corridor_Line_Flow']
        sec['_corridor_src'] = 'Corridor_Line_Flow 시트'
    elif 'Line_Flow' in raw and 'Line_Info' in raw:
        sec['corridor'] = build_corridor(raw['Line_Flow'], raw['Line_Info'])
        sec['_corridor_src'] = 'Line_Flow + Line_Info 에서 계산'
    pairs = [('gen', 'National_Gen_Output', None), ('link', 'National_Link_Flow', 'Link_Flow'),
             ('store_e', 'National_Storage_Energy', None), ('store_p', 'National_Storage_Power', 'Storage_Power'),
             ('curtail', 'RE_Curtailment_Timeseries', None), ('loads', 'National_Loads', None)]
    for key, sheet, fallback in pairs:
        if sheet in raw:
            sec[key] = raw[sheet]
        elif fallback and fallback in raw:
            sec[key] = _national(raw[fallback])
    return sec


def _line_utilization(raw):
    lf, li = raw.get('Line_Flow'), raw.get('Line_Info')
    if lf is None or li is None:
        return None
    cap = li['s_nom_opt'] if 's_nom_opt' in li else li['s_nom']
    cols = [c for c in lf.columns if c in cap.index and float(cap[c]) > 0]
    if not cols:
        return None
    return (lf[cols].abs() / cap[cols].astype(float)).max(axis=1)


# ── 공개 API ─────────────────────────────────────────────────────────
def load_result(path, use_cache=True, progress=None):
    """결과 파일 → {'sections', 'balance', 'line_util', 'source_sheets', ...}"""
    path = os.path.abspath(path)
    st = os.stat(path)
    cache_dir = os.path.join(os.path.dirname(path), CACHE_DIRNAME)
    cache = os.path.join(cache_dir, os.path.basename(path) + '.pkl')
    sig = (CACHE_VERSION, st.st_size, int(st.st_mtime))
    if use_cache and os.path.exists(cache):
        try:
            with open(cache, 'rb') as f:
                obj = pickle.load(f)
            if obj.get('_sig') == sig:
                if progress:
                    progress('캐시에서 불러옴')
                return obj
        except Exception:
            pass

    if path.lower().endswith('.nc'):
        raw = _raw_from_network(path, progress)
    else:
        raw = _read_sheets(path, RAW_SHEETS, progress)
    if 'National_Gen_Output' not in raw and 'Generator_Output' in raw['_sheetnames']:
        go = _read_sheets(path, ['Generator_Output'], progress).get('Generator_Output')
        if go is not None and go.shape[1] > 0:
            raw['National_Gen_Output'] = _national(go)
    if 'National_Gen_Output' not in raw:
        raise ValueError('시간별 발전 결과(National_Gen_Output / Generator_Output)가 없는 파일입니다. '
                         'PyPSA 결과 파일인지, 최적화 결과가 정상 저장되었는지 확인하세요.')
    if progress:
        progress('수급 데이터 계산 중')
    obj = {
        '_sig': sig,
        'path': path,
        'sections': build_sections(raw),
        'balance': build_el_balance(raw),
        'line_util': _line_utilization(raw),
        'source_sheets': raw['_sheetnames'],
    }
    obj['index'] = obj['balance']['supply'].index
    try:
        os.makedirs(cache_dir, exist_ok=True)
        with open(cache, 'wb') as f:
            pickle.dump(obj, f, protocol=pickle.HIGHEST_PROTOCOL)
    except Exception:
        pass
    return obj


def list_result_files(folder, recursive=False):
    pat = re.compile(r'\.(xlsx|xlsm|nc)$', re.I)
    out = []
    if recursive:
        for root, dirs, files in os.walk(folder):
            dirs[:] = [d for d in dirs if d != CACHE_DIRNAME and not d.startswith('.')]
            out += [os.path.join(root, f) for f in files if pat.search(f) and not f.startswith('~$')]
    else:
        out = [os.path.join(folder, f) for f in os.listdir(folder)
               if pat.search(f) and not f.startswith('~$') and os.path.isfile(os.path.join(folder, f))]
    out = [p for p in out if not os.path.basename(p).startswith('AW_')]
    return sorted(out, key=lambda p: os.path.getmtime(p), reverse=True)


def window(obj, start, hours):
    """start(Timestamp) 이후 hours 개 스냅샷 위치 slice."""
    idx = obj['index']
    i0 = int(idx.searchsorted(pd.Timestamp(start)))
    i0 = min(max(i0, 0), max(len(idx) - 1, 0))
    return slice(i0, min(i0 + int(hours), len(idx)))


def auto_start(obj, mode, hours):
    idx = obj['index']
    b = obj['balance']
    if mode == 'line' and obj.get('line_util') is not None:
        s = obj['line_util']
    elif mode == 'curtail' and 'Curtail' in b['consume']:
        s = b['consume']['Curtail']
    elif mode == 'demand':
        s = b['demand']
    elif mode == 'net_low':
        s = -(b['demand'] - b['supply'].reindex(columns=['PV', 'WT']).fillna(0).sum(axis=1))
    else:
        return idx[0]
    loc = idx.get_loc(s.rolling(min(hours, len(s)), min_periods=1).sum().shift(-(hours - 1)).fillna(-np.inf).idxmax())
    return idx[max(0, min(loc, len(idx) - hours))]


def balance_table(obj, sl):
    b = obj['balance']
    lab_s = {k: l for k, l, _ in SUPPLY_DEF}
    lab_c = {k: l for k, l, _ in CONSUME_DEF}
    sup = b['supply'].iloc[sl].rename(columns=lambda k: f'[공급] {lab_s.get(k, k)}')
    con = b['consume'].iloc[sl].rename(columns=lambda k: f'[소비] {lab_c.get(k, k)}')
    dem = b['demand'].iloc[sl].rename(f'[소비] {DEMAND_LABEL}')
    df = pd.concat([sup, dem, con], axis=1)
    df['수급차(공급-소비)'] = b['residual'].iloc[sl]
    df.index.name = 'Timestamp'
    return df
