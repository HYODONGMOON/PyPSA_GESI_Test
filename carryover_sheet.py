"""다음 연도 승계 입력표.

최적화 결과에서 확장 가능 설비의 최적 용량을 뽑아, 다음 연도 분석 때 interface.xlsx의
어느 셀에 어떤 값을 넣어야 하는지 정리한다. 지역 시트(지역_XXX)의 용량 셀 수식을 따라가
실제 입력 셀(대부분 '인터페이스' 시트)을 찾는다.

사용:
  - save_results()에서 write_carryover_sheet(writer, network, interface_path) 호출
  - 기존 결과에 대해: python carryover_sheet.py <result.nc> [--xlsx 결과.xlsx | --out 파일.xlsx]
"""
import os
import re
import shutil
import sys
import tempfile

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter, column_index_from_string

SHEET_NAME = '다음연도_승계입력'
REGION_KR = {'BSN': '부산울산', 'CBD': '충북', 'CND': '대전세종충남', 'GBD': '대구경북', 'GGD': '서울경기',
             'GND': '경남', 'GWD': '강원', 'ICN': '인천', 'JBD': '전북', 'JJD': '제주', 'JND': '광주전남'}
KIND_KR = {'PV': '태양광', 'WT': '풍력', 'ESS': 'ESS 저장용량', 'H2_Storage': '수소 저장', 'H_Storage': '열 저장',
           'EV_DSM': 'EV DSM', 'Electrolyser': '수전해', 'ESS_Ch': 'ESS 충전', 'ESS_Disch': 'ESS 방전',
           'Fuelcell': '연료전지', 'PV_RE': 'RE100 태양광', 'WT_RE': 'RE100 풍력', 'ESS_RE': 'RE100 ESS',
           'EL_to_DC_Fab_RE': 'RE100 연계선'}
SECTIONS = {'발전기': ('generators', 'p_nom', 'D', 'MW'),
            '저장장치': ('stores', 'e_nom', 'D', 'MWh'),
            '링크': ('links', 'p_nom', 'J', 'MW')}
IFACE = '인터페이스'
CUR_COLS = [get_column_letter(i) for i in range(column_index_from_string('D'), column_index_from_string('N') + 1)]
PREV_COLS = [get_column_letter(i) for i in range(column_index_from_string('T'), column_index_from_string('AD') + 1)]
REF_RE = re.compile(r"^=\s*'?([^'!]+?)'?!\$?([A-Z]{1,3})\$?(\d+)\s*$")

def _text(v):
    return getattr(v, 'text', v)


def _display(content):
    """결과 시트에 수식으로 기록되지 않도록 '=' 없이 표시한다."""
    if content is None:
        return ''
    s = str(content)
    return f'수식: {s[1:]}' if s.startswith('=') else f'값: {s}'


def _grid(ws, max_col=None):
    g = {}
    for r, row in enumerate(ws.iter_rows(max_col=max_col), start=1):
        for c, cell in enumerate(row, start=1):
            v = getattr(cell, 'value', None)
            if v is not None:
                g[f'{get_column_letter(c)}{r}'] = v
    return g


def _region_capacity_cells(grid_f, region):
    """지역 시트에서 설비별 용량 셀 위치와 내용을 찾는다: {(comp, 네트워크 이름): (셀, 내용, 단위, 행)}"""
    out, sec = {}, None
    rows = sorted({int(re.sub(r'^[A-Z]+', '', k)) for k in grid_f})
    for r in rows:
        a = grid_f.get(f'A{r}')
        if not isinstance(a, str):
            continue
        key = a.strip()
        if key in SECTIONS:
            sec = key
            continue
        if sec is None or key in ('이름', ''):
            continue
        comp, _, col, unit = SECTIONS[sec]
        out[(comp, f'{region}_{key}')] = (f'{col}{r}', _text(grid_f.get(f'{col}{r}')), unit, r, key)
    return out


def _block_label(col, year):
    nxt = year + 5 if isinstance(year, (int, float)) else None
    if col in CUR_COLS:
        return f'해당연도 블록 D~N열 (B1={nxt}이면 {nxt}년 값 자리)' if nxt else '해당연도 블록 D~N열'
    if col in PREV_COLS:
        return f'직전연도 블록 T~AD열 (B1={nxt}이면 {year}년 값 자리)' if nxt else '직전연도 블록 T~AD열'
    return ''


def build_carryover_table(network, interface_path, year=None):
    """확장 가능 설비의 최적 용량을 interface 입력 셀 단위로 정리한 표와 기준 연도를 반환."""
    wb_f = openpyxl.load_workbook(interface_path, read_only=True, data_only=False)
    wb_v = openpyxl.load_workbook(interface_path, read_only=True, data_only=True)
    try:
        iface_f, iface_v = _grid(wb_f[IFACE]), _grid(wb_v[IFACE])
        year = year if year is not None else iface_v.get('B1')
        regions = sorted({str(b).split('_')[0] for b in network.buses.index} & set(REGION_KR))
        rows = []
        for reg in regions:
            sn = f'지역_{reg}'
            if sn not in wb_f.sheetnames:
                continue
            reg_f = _grid(wb_f[sn], max_col=20)
            for (comp, nm), (cell, content, unit, r, kind) in _region_capacity_cells(reg_f, reg).items():
                df = getattr(network, comp)
                attr = SECTIONS[[k for k, v in SECTIONS.items() if v[0] == comp][0]][1]
                if nm not in df.index or 'slack' in nm.lower() or not bool(df.at[nm, f'{attr}_extendable']):
                    continue
                opt_col = f'{attr}_opt'
                new = float(df.at[nm, opt_col]) if opt_col in df.columns and pd.notna(df.at[nm, opt_col]) else float(df.at[nm, attr])
                existing = float(df.at[nm, attr])
                if max(abs(new), abs(existing)) < 1e-3:
                    continue
                m = REF_RE.match(str(content)) if isinstance(content, str) else None
                if m and m.group(1) in wb_f.sheetnames:
                    t_sheet, t_cell = m.group(1), f'{m.group(2)}{m.group(3)}'
                    t_grid_f = iface_f if t_sheet == IFACE else _grid(wb_f[t_sheet])
                    t_content = _text(t_grid_f.get(t_cell))
                    item = iface_v.get(f'C{m.group(3)}') if t_sheet == IFACE else None
                    block = _block_label(m.group(2), year) if t_sheet == IFACE else ''
                else:
                    t_sheet, t_cell, t_content = sn, cell, content
                    item, block = None, '지역 시트 직접 입력'
                note = []
                if isinstance(t_content, str) and t_content.startswith('='):
                    note.append('수식 셀: 값으로 덮어쓰기')
                if kind == 'Fuelcell':
                    note.append('최대용량(지역 시트 M열 = 기존/0.25)도 함께 바뀜')
                max_attr = f'{attr}_max'
                mx = float(df.at[nm, max_attr]) if max_attr in df.columns else np.inf
                if np.isfinite(mx) and new > mx + 1:
                    note.append(f'현재 상한 {mx:,.0f} 초과(모형이 상한을 기존 용량으로 올림)')
                rows.append({
                    '지역': reg, '지역명': REGION_KR.get(reg, reg), '설비': KIND_KR.get(kind, kind), '네트워크 이름': nm,
                    '단위': unit, '입력 시트': t_sheet, '입력 셀': t_cell,
                    '승계값 (입력할 값)': new, '이번 분석 기존 설비': existing, '이번 분석 신규 증설': new - existing,
                    'interface 항목 (C열)': item, '블록': block,
                    '현재 셀 내용': _display(t_content), '지역 시트 참조': f'{sn}!{cell}',
                    '비고': '; '.join(note)})
        table = pd.DataFrame(rows)
        if not table.empty:
            order = {k: i for i, k in enumerate(KIND_KR.values())}
            table = table.sort_values(['설비', '지역'], key=lambda s: s.map(order) if s.name == '설비' else s).reset_index(drop=True)
        return table, year
    finally:
        wb_f.close(); wb_v.close()


def _summary(table):
    if table.empty:
        return pd.DataFrame()
    s = table.groupby(['설비', '단위'], sort=False)[['이번 분석 기존 설비', '이번 분석 신규 증설', '승계값 (입력할 값)']].sum() / 1e3
    s.columns = ['기존 설비 합계 (GW/GWh)', '신규 증설 합계 (GW/GWh)', '승계값 합계 (GW/GWh)']
    s['입력 셀 수'] = table.groupby(['설비', '단위'], sort=False).size()
    s = s.reset_index()
    s['단위'] = s['단위'].map({'MW': 'GW', 'MWh': 'GWh'}).fillna(s['단위'])
    return s.rename(columns={c: c.replace(' (GW/GWh)', '') for c in s.columns})


def _info_lines(year):
    nxt = year + 5 if isinstance(year, (int, float)) else '다음 연도'
    return [
        f'이번 분석 연도 (인터페이스!B1): {year}  →  승계 대상 연도: {nxt}',
        f"입력 방법: interface.xlsx의 '인터페이스'!B1을 {nxt}로 바꾼 뒤, 아래 표의 '입력 셀'에 '승계값'을 값으로 입력한다.",
        '대상: 최적화로 용량이 정해지는 확장 가능 설비만. 원자력·LNG·석탄·HP·EV 충전기·HVDC 등 비확장 설비와 수요·CO2 제약은 다음 연도 interface 값을 그대로 쓴다.',
        '승계 방식: 이번 결과의 최적 용량으로 대체(기존 설비로 입력). 모형이 기존 용량을 확장 하한으로 적용하므로 다음 연도 비용에는 추가분만 반영된다. 설비 수명에 따른 폐지는 반영하지 않는다.',
        "주의: '직전연도 블록'(T~AD열) 셀은 B1-5년 계획값을 불러오는 배열수식이다. 값으로 덮어쓰면 수식이 사라지므로 interface 원본을 따로 보관해 둘 것.",
        '열보조공급(H_Fallback)·LNG 연료공급·slack 등 모형이 자동으로 만드는 보조 설비는 승계하지 않는다.',
    ]


def write_carryover_sheet(writer, network, interface_path, year=None):
    """열려 있는 pd.ExcelWriter(openpyxl)에 승계 입력 시트를 추가한다."""
    table, year = build_carryover_table(network, interface_path, year)
    info = pd.DataFrame({'안내': _info_lines(year)})
    info.to_excel(writer, sheet_name=SHEET_NAME, index=False, startrow=0)
    summ = _summary(table)
    r = len(info) + 2
    summ.to_excel(writer, sheet_name=SHEET_NAME, index=False, startrow=r)
    r += len(summ) + 3
    table.to_excel(writer, sheet_name=SHEET_NAME, index=False, startrow=r)
    ws = writer.sheets[SHEET_NAME]
    widths = {'A': 14, 'B': 12, 'C': 14, 'D': 22, 'E': 7, 'F': 12, 'G': 9, 'H': 14, 'I': 16, 'J': 12,
              'K': 9, 'L': 14, 'M': 14, 'N': 28, 'O': 40, 'P': 18, 'Q': 50}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=r + 2, max_row=r + 1 + len(table), min_col=8, max_col=10):
        for c in row:
            c.number_format = '#,##0.0'
    for row in ws.iter_rows(min_row=len(info) + 4, max_row=len(info) + 3 + len(summ), min_col=3, max_col=5):
        for c in row:
            c.number_format = '#,##0.00'
    print(f"{SHEET_NAME} 시트 저장 완료: 확장 설비 입력 셀 {len(table)}개 (기준 연도 {year})")
    return table


def _load_network(nc_path):
    import pypsa
    try:
        return pypsa.Network(nc_path)
    except Exception:
        tmp = os.path.join(tempfile.gettempdir(), 'carryover_tmp.nc')
        shutil.copyfile(nc_path, tmp)
        return pypsa.Network(tmp)


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description='결과 네트워크(.nc)로 다음 연도 승계 입력 시트를 만든다.')
    ap.add_argument('nc')
    ap.add_argument('--xlsx', help='이 결과 xlsx에 시트를 추가한다(기존 같은 이름 시트는 교체)')
    ap.add_argument('--out', help='별도 xlsx 파일로 저장 (기본: nc 옆 <이름>_승계입력.xlsx)')
    ap.add_argument('--year', type=int, help='결과의 분석 연도 (기본: interface 인터페이스!B1)')
    ap.add_argument('--interface', default=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'interface.xlsx'))
    a = ap.parse_args(argv)
    n = _load_network(a.nc)
    if a.xlsx:
        with pd.ExcelWriter(a.xlsx, engine='openpyxl', mode='a', if_sheet_exists='replace') as w:
            write_carryover_sheet(w, n, a.interface, a.year)
        print(f'추가: {a.xlsx}')
    else:
        out = a.out or os.path.splitext(a.nc)[0] + '_승계입력.xlsx'
        with pd.ExcelWriter(out, engine='openpyxl') as w:
            write_carryover_sheet(w, n, a.interface, a.year)
        print(f'저장: {out}')


if __name__ == '__main__':
    main(sys.argv[1:])
