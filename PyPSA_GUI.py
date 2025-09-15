import os as _os
import sys as _sys

def _ensure_proj_lib_env():
    try:
        if 'PROJ_LIB' in _os.environ and _os.environ.get('PROJ_LIB'):
            return
        candidates = []
        try:
            # Conda/Windows 일반 경로
            candidates.append(_os.path.join(_sys.prefix, 'Library', 'share', 'proj'))
            candidates.append(_os.path.join(_sys.prefix, 'share', 'proj'))
        except Exception:
            pass
        # 고정 경로(사용 환경)
        candidates.append(r'C:\ProgramData\anaconda3\envs\pypsa_env\Library\share\proj')
        for c in candidates:
            try:
                if c and _os.path.exists(c):
                    _os.environ['PROJ_LIB'] = c
                    break
            except Exception:
                continue
    except Exception:
        pass

_ensure_proj_lib_env()
import pypsa
import pandas as pd
import numpy as np
from datetime import datetime
import os
import traceback
import sys
import copy


# src 폴더를 Python 경로에 추가
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

# 지도 모듈은 선택적 임포트로 처리하여 rasterio/GDAL 미설치 시에도 실행이 가능하도록 함
try:
    from korea_map import KoreaMapVisualizer  # noqa: F401
except Exception as _map_e:
    KoreaMapVisualizer = None
    try:
        _force_print(f"지도 모듈 로드 경고: {str(_map_e)}")
    except Exception:
        pass

# 상수 정의
INPUT_FILE = "integrated_input_data.xlsx"

def _normalize_region_code(value):
    try:
        s = str(value).strip()
        if not s:
            return s
        codes = ['SEL','BSN','DGU','ICN','GWJ','DJN','USN','SJG','GGD','GWD','CBD','CND','JBD','JND','GBD','GND','JJD']
        if s.upper() in codes:
            return s.upper()
        name_to_code = {
            '서울특별시':'SEL','부산광역시':'BSN','대구광역시':'DGU','인천광역시':'ICN','광주광역시':'GWJ',
            '대전광역시':'DJN','울산광역시':'USN','세종특별자치시':'SJG','경기도':'GGD','강원도':'GWD',
            '충청북도':'CBD','충청남도':'CND','전라북도':'JBD','전라남도':'JND','경상북도':'GBD','경상남도':'GND','제주특별자치도':'JJD'
        }
        return name_to_code.get(s, s)
    except Exception:
        return str(value)

def read_input_data(input_file):
    """Excel 파일에서 입력 데이터 읽기"""
    try:
        # 파일 경로 및 수정 시간 로깅
        root_dir = os.path.dirname(__file__) 
        integrated_path = os.path.abspath(os.path.join(root_dir, input_file))
        interface_path = os.path.abspath(os.path.join(root_dir, 'interface.xlsx'))
        try:
            if os.path.exists(integrated_path):
                print(f"integrated_input_data 경로: {integrated_path}")
                print(f"integrated_input_data 수정 시간: {datetime.fromtimestamp(os.path.getmtime(integrated_path))}")
            if os.path.exists(interface_path):
                print(f"interface.xlsx 경로: {interface_path}")
                print(f"interface.xlsx 수정 시간: {datetime.fromtimestamp(os.path.getmtime(interface_path))}")
        except Exception:
            pass

        xls = pd.ExcelFile(input_file)
        input_data = {
            'buses': pd.DataFrame(),
            'generators': pd.DataFrame(),
            'loads': pd.DataFrame(),
            'stores': pd.DataFrame(),
            'links': pd.DataFrame(),
            'lines': pd.DataFrame(),
            'timeseries': pd.DataFrame(),
            'renewable_patterns': pd.DataFrame(),
            'load_patterns': pd.DataFrame()
        }
        
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            # 컬럼명이 문자열인 경우에만 strip 적용
            if df.columns.dtype == 'object':
                try:
                    df.columns = df.columns.str.strip()
                except:
                    # 문자열이 아닌 컬럼명이 있는 경우 문자열로 변환 후 strip
                    df.columns = [str(col).strip() for col in df.columns]
            input_data[sheet_name] = df
        
        # Fallback: 통합 파일에 패턴 시트가 없으면 interface.xlsx에서 보강
        try:
            if (('load_patterns' not in xls.sheet_names) or input_data['load_patterns'].empty) and os.path.exists(interface_path):
                lp = pd.read_excel(interface_path, sheet_name='load_patterns')
                if lp is not None and not lp.empty:
                    input_data['load_patterns'] = lp
                    print("load_patterns를 interface.xlsx에서 보강 로딩했습니다.")
        except Exception as e:
            print(f"load_patterns 보강 로딩 실패: {str(e)}")
        try:
            if (('renewable_patterns' not in xls.sheet_names) or input_data['renewable_patterns'].empty) and os.path.exists(interface_path):
                rp = pd.read_excel(interface_path, sheet_name='renewable_patterns')
                if rp is not None and not rp.empty:
                    input_data['renewable_patterns'] = rp
                    print("renewable_patterns를 interface.xlsx에서 보강 로딩했습니다.")
        except Exception as e:
            print(f"renewable_patterns 보강 로딩 실패: {str(e)}")
        try:
            if os.path.exists(interface_path):
                sd = pd.read_excel(interface_path, sheet_name='시나리오_에너지수요')
                if sd is not None and not sd.empty:
                    input_data['시나리오_에너지수요'] = sd
                    print("시나리오_에너지수요를 interface.xlsx에서 로딩했습니다.")
        except Exception as e:
            print(f"시나리오_에너지수요 로딩 실패: {str(e)}")

        # Fallback: lines 시트가 없거나 비어있으면 interface.xlsx의 '지역간 연결'에서 생성
        try:
            if (('lines' not in xls.sheet_names) or input_data['lines'].empty) and os.path.exists(interface_path):
                fb = _fallback_build_lines_from_interface(input_data, interface_path)
                if fb is not None and not fb.empty:
                    input_data['lines'] = fb
                    print("lines를 interface.xlsx의 '지역간 연결'에서 폴백 생성했습니다.")
                    # 통합 파일에도 저장해 다음 실행부터 바로 사용
                    _persist_lines_to_integrated(integrated_path, fb)
        except Exception as e:
            print(f"lines 폴백 생성 실패: {str(e)}")
        
        return input_data
        
    except Exception as e:
        print(f"데이터 읽기 오류: {str(e)}")
        raise

def adjust_pattern_length(pattern_values, required_length):
    """패턴 길이를 필요한 길이에 맞게 조정"""
    if len(pattern_values) == required_length:
        return pattern_values
    
    # 패턴이 더 짧은 경우
    elif len(pattern_values) < required_length:
        # 부족한 만큼 처음부터 반복
        repetitions = required_length // len(pattern_values) + 1
        extended_pattern = np.tile(pattern_values, repetitions)
        return extended_pattern[:required_length]
    
    # 패턴이 더 긴 경우
    else:
        return pattern_values[:required_length]

def normalize_pattern(pattern):
    """발전 패턴을 0~1 사이로 정규화"""
    if np.max(pattern) > 0:
        return pattern / np.max(pattern)
    return pattern

def _get_scenario_year(input_data):
    try:
        if 'timeseries' in input_data and not input_data['timeseries'].empty:
            start = pd.to_datetime(input_data['timeseries'].iloc[0]['start_time'])
            return int(start.year)
    except Exception:
        pass
    return None

def _parse_scenario_demand(input_data, scenario_year):
    key_map = {
        'year': ['year', '연도', '년도'],
        'region': ['region', '지역', '코드', '지역코드'],
        'EL': ['EL', '전력', '전력(MWh)', '전력_MWh'],
        'H': ['H', '열', '열(MWh)', '열_MWh', 'HEAT'],
        'H2': ['H2', '수소', '수소(MWh)', '수소_MWh']
    }
    def find_col(cols, candidates):
        for cand in candidates:
            for c in cols:
                if str(c).strip().lower() == str(cand).strip().lower():
                    return c
        return None
    scenario = {}
    if '시나리오_에너지수요' not in input_data:
        return scenario
    df = input_data['시나리오_에너지수요']
    if df is None or df.empty:
        return scenario
    year_col = find_col(df.columns, key_map['year'])
    region_col = find_col(df.columns, key_map['region'])
    el_col = find_col(df.columns, key_map['EL'])
    h_col = find_col(df.columns, key_map['H'])
    h2_col = find_col(df.columns, key_map['H2'])
    if year_col is None or region_col is None:
        return scenario
    try:
        filt = df[year_col].astype(int) == int(scenario_year)
        ydf = df.loc[filt].copy()
    except Exception:
        ydf = df.copy()
    for _, row in ydf.iterrows():
        raw_region = str(row[region_col]).strip()
        if '_' in raw_region:
            raw_region = raw_region.split('_')[0]
        region = _normalize_region_code(raw_region)
        if not region or str(region).lower() == 'nan':
            continue
        if el_col in ydf.columns and pd.notna(row.get(el_col)):
            scenario[(region, 'EL')] = float(row[el_col])
        if h_col in ydf.columns and pd.notna(row.get(h_col)):
            scenario[(region, 'H')] = float(row[h_col])
        if h2_col in ydf.columns and pd.notna(row.get(h2_col)):
            scenario[(region, 'H2')] = float(row[h2_col])
    return scenario

def _apply_scenario_demand_scaling(network, input_data):
    try:
        # 멀티년 분석 등에서 시나리오 스케일링 비활성화 플래그
        if os.environ.get('DISABLE_SCALING') == '1':
            print("시나리오 수요 스케일링 비활성화됨(DISABLE_SCALING=1)")
            return
        scenario_year = _get_scenario_year(input_data)
        if scenario_year is None:
            return
        scenario = _parse_scenario_demand(input_data, scenario_year)
        if not scenario:
            print("시나리오_에너지수요 시트가 없거나 해당 연도 데이터가 없습니다. 기본 부하를 사용합니다.")
            return
        group_to_loads = {}
        for load_name in network.loads.index:
            region = load_name.split('_')[0] if '_' in load_name else None
            demand_type = None
            lname = load_name.lower()
            if '_demand_el' in lname:
                demand_type = 'EL'
            elif '_demand_h2' in lname:
                demand_type = 'H2'
            elif '_demand_h' in lname:
                demand_type = 'H'
            if region and demand_type:
                key = (region, demand_type)
                group_to_loads.setdefault(key, []).append(load_name)
        total_scaled_groups = 0
        for key, names in group_to_loads.items():
            region, dtype = key
            target = scenario.get((region, dtype))
            if target is None:
                continue
            current = 0.0
            for nm in names:
                if nm in network.loads_t.p_set.columns:
                    current += float(np.nansum(network.loads_t.p_set[nm].values))
            if current <= 0:
                if target > 0:
                    print(f"경고: {region}-{dtype} 현재 부하 합계가 0입니다. 스케일링을 건너뜁니다.")
                continue
            # 이미 패턴 합계=1로 연간 총량을 분배했으므로, 스케일은 미세 오차 보정 수준이어야 함
            scale = target / current
            for nm in names:
                if nm in network.loads_t.p_set.columns:
                    network.loads_t.p_set[nm] = network.loads_t.p_set[nm] * scale
            print(f"수요 스케일링 적용: {region}-{dtype}: 현재 {current:.1f} → 목표 {target:.1f} (배율 {scale:.6f})")
            total_scaled_groups += 1
        if total_scaled_groups == 0:
            print("시나리오 스케일링 대상 그룹이 없습니다. 이름 규칙 또는 시트 컬럼을 확인하세요.")
    except Exception as e:
        print(f"수요 시나리오 스케일링 중 오류: {str(e)}")

def _get_load_pattern(input_data, region, demand_type, snapshots_len):
    try:
        if 'load_patterns' not in input_data or input_data['load_patterns'].empty:
            return None
        df = input_data['load_patterns'].copy()

        # 1) 고정 위치(B8:D8765)에서 직접 읽기 시도
        try:
            base_len = 8760
            start_row = 7  # Excel 8행 → iloc 7
            col_map = {'EL': 1, 'H': 2, 'H2': 3}  # B,C,D
            if demand_type in col_map and df.shape[0] > start_row + 10 and df.shape[1] > col_map[demand_type]:
                series = pd.to_numeric(df.iloc[start_row:, col_map[demand_type]], errors='coerce')
                values = series.dropna().astype(float).values
                if len(values) > 0:
                    # 기본 8760 채움: 부족 시 타일링
                    if len(values) >= base_len:
                        pattern_base = values[:base_len]
                    else:
                        repeats = base_len // len(values) + 1
                        pattern_base = np.tile(values, repeats)[:base_len]
                    # 스냅샷 길이에 맞추어 1-based 타일링(8761→1...)
                    if snapshots_len <= base_len:
                        pattern_full = pattern_base[:snapshots_len]
                    else:
                        repeats = snapshots_len // base_len + 1
                        pattern_full = np.tile(pattern_base, repeats)[:snapshots_len]
                    # 디버그 로그
                    try:
                        label = {'EL': '전력(B열)', 'H': '열(C열)', 'H2': '수소(D열)'}[demand_type]
                        sample_n = min(8, len(pattern_base))
                        print(f"load_patterns 고정위치 사용: {label}, 시작행 Excel 8행 기준")
                        print(f"패턴 길이(기본): {len(pattern_base)} (최대 {np.nanmax(pattern_base):.6f}, 평균 {np.nanmean(pattern_base):.6f})")
                        print(f"패턴 샘플 앞 {sample_n}개: {np.round(pattern_base[:sample_n], 6).tolist()}")
                        if len(pattern_base) >= 24:
                            day_slice = pattern_base[:24]
                            print(f"하루(1~24시) 최소/최대: {np.min(day_slice):.6f}/{np.max(day_slice):.6f}")
                    except Exception:
                        pass
                    return pattern_full
        except Exception as e:
            print(f"고정 위치 패턴 로딩 실패(폴백 사용): {str(e)}")

        # 2) 폴백: 기존 컬럼 매칭 로직
        # 컬럼 정리: 공백 제거 및 한국어 기본 패턴명을 'pattern'으로 표준화
        cleaned_cols = []
        for c in df.columns:
            name = str(c).strip()
            if ('부하' in name and '패턴' in name) or name.lower() in ['default', 'pattern']:
                name = 'pattern'
            cleaned_cols.append(name)
        df.columns = cleaned_cols

        # hour/시간 컬럼 확인(이름 부분일치 + 수치 램프 감지)
        hour_col = None
        # 1) 이름 부분일치로 우선 탐색
        for c in df.columns:
            cl = str(c).strip().lower()
            if ('hour' in cl) or ('시간' in cl) or (cl in ['h', 't', 'index']):
                hour_col = c
                break
        # 2) 이름으로 못 찾으면, 값 패턴이 1..8760 류의 정수 램프인지 검사하여 선택
        if hour_col is None:
            for c in df.columns:
                s = pd.to_numeric(df[c], errors='coerce')
                if s.notna().sum() < 10:
                    continue
                vals = s.dropna().values
                # 정수성, 범위, 단조 증가(많은 구간) 검사
                is_int_like = np.nanmax(np.abs(vals - np.round(vals))) < 1e-6
                if not is_int_like:
                    continue
                vmin, vmax = np.nanmin(vals), np.nanmax(vals)
                if vmin >= 1 and vmax <= 8760 and vmax - vmin >= 1000:
                    diffs = np.diff(vals[:min(len(vals), 2000)])
                    if np.mean(diffs >= 0) > 0.95:  # 거의 단조 증가
                        hour_col = c
                        break

        # 후보 컬럼 선택
        candidates_by_type = {
            'EL': [f"{region}_EL", f"{region}_electricity", 'EL', 'electricity', '전력'],
            'H': [f"{region}_H", f"{region}_heating", 'H', 'heat', 'heating', '열'],
            'H2': [f"{region}_H2", f"{region}_hydrogen", 'H2', 'hydrogen', '수소']
        }
        candidates = candidates_by_type.get(demand_type, [])
        # 지역 무관 전역 후보 추가
        if demand_type == 'EL':
            candidates += ['EL', 'electricity', '전력']
        elif demand_type == 'H':
            candidates += ['H', 'heat', 'heating', '열']
        elif demand_type == 'H2':
            candidates += ['H2', 'hydrogen', '수소']
        # 기본 패턴 컬럼
        candidates += ['pattern']

        col = None
        for cand in candidates:
            if cand in df.columns:
                col = cand
                break

        # 마지막 fallback: 첫 번째 수치형 컬럼 탐색(hour/시간 유사 컬럼 제외)
        if col is None:
            numeric_cols = []
            for c in df.columns:
                cl = str(c).strip().lower()
                if hour_col and c == hour_col:
                    continue
                if ('hour' in cl) or ('시간' in cl) or (cl in ['h', 't', 'index']):
                    continue
                s = pd.to_numeric(df[c], errors='coerce')
                if s.notna().sum() > 0:
                    # 값 자체가 1..8760 류 램프인 경우 제외(시간 컬럼 가능성 높음)
                    vals = s.dropna().values
                    if len(vals) >= 100 and np.nanmax(vals) <= 100000:
                        diffs = np.diff(vals[:min(len(vals), 2000)])
                        ramp_like = np.mean(diffs >= 0) > 0.95 and np.nanmin(vals) >= 1
                        if ramp_like:
                            continue
                    numeric_cols.append(c)
            if numeric_cols:
                col = numeric_cols[0]
            else:
                return None

        # 기본 길이는 8760으로 간주(윤년/추가 스냅샷은 타일링)
        base_len = 8760

        if hour_col:
            # 1-based 매핑: 시간값 1..8760 → 인덱스 0..8759
            base = np.zeros(base_len, dtype=float)
            series = pd.to_numeric(df[col], errors='coerce')
            hours = pd.to_numeric(df[hour_col], errors='coerce')
            for v, h in zip(series, hours):
                if pd.notna(v) and pd.notna(h):
                    try:
                        h_int = int(float(h))
                    except Exception:
                        continue
                    if 1 <= h_int <= base_len:
                        base[h_int - 1] = float(v)
            pattern_base = base
        else:
            # 시간 컬럼이 없으면, 수치값만 추출하여 8760 길이로 보정(부족 시 타일링)
            series = pd.to_numeric(df[col], errors='coerce')
            values = series.dropna().astype(float).values
            if len(values) == 0:
                return None
            if len(values) >= base_len:
                pattern_base = values[:base_len]
            else:
                repeats = base_len // len(values) + 1
                pattern_base = np.tile(values, repeats)[:base_len]

        # 스냅샷 길이에 맞추어 1-based 타일링(8761→1...)
        if snapshots_len <= base_len:
            pattern_full = pattern_base[:snapshots_len]
        else:
            repeats = snapshots_len // base_len + 1
            pattern_full = np.tile(pattern_base, repeats)[:snapshots_len]

        # 디버그 로그: 선택된 컬럼과 패턴 샘플/단조 증가 여부
        try:
            print(f"load_patterns 컬럼: {list(df.columns)[:10]}{'...' if len(df.columns)>10 else ''}")
            print(f"선택된 시간 컬럼: {hour_col}")
            print(f"선택된 패턴 컬럼: {col}")
            sample_n = min(8, len(pattern_base))
            print(f"패턴 샘플(기본 기준 앞 {sample_n}개): {np.round(pattern_base[:sample_n], 6).tolist()}")
            if len(pattern_base) >= 24:
                day_slice = pattern_base[:24]
                print(f"하루(1~24시) 최소/최대: {np.min(day_slice):.6f}/{np.max(day_slice):.6f}")
        except Exception:
            pass

        return pattern_full
    except Exception:
        return None

def create_network(input_data):
    try:
        network = pypsa.Network()
        
        # carriers 정의
        carriers = {
            'AC': {'name': 'AC', 'co2_emissions': 0},
            'DC': {'name': 'DC', 'co2_emissions': 0},
            'electricity': {'name': 'electricity', 'co2_emissions': 0},
            'coal': {'name': 'coal', 'co2_emissions': 0.9},
            'gas': {'name': 'gas', 'co2_emissions': 0.4},
            'nuclear': {'name': 'nuclear', 'co2_emissions': 0},
            'solar': {'name': 'solar', 'co2_emissions': 0},
            'wind': {'name': 'wind', 'co2_emissions': 0},
            'hydrogen': {'name': 'hydrogen', 'co2_emissions': 0},
            'heat': {'name': 'heat', 'co2_emissions': 0}
        }
        
        # carriers 추가
        for carrier, specs in carriers.items():
            network.add("Carrier",
                       name=specs['name'],
                       co2_emissions=specs['co2_emissions'])
        
        # 시간 설정
        if 'timeseries' in input_data and not input_data['timeseries'].empty:
            ts = input_data['timeseries'].iloc[0]
            snapshots = pd.date_range(
                start=ts['start_time'],
                end=ts['end_time'],
                freq=ts['frequency'],
                inclusive='left'
            )
            network.set_snapshots(snapshots)
            snapshots_length = len(snapshots)
        else:
            # 기본 시간 설정 (2024년 1년간, 1시간 간격)
            print("⚠️ timeseries 시트가 없거나 비어있습니다. 기본 시간 설정을 사용합니다.")
            snapshots = pd.date_range(
                start='2024-01-01 00:00:00',
                end='2025-01-01 00:00:00',
                freq='1h',
                inclusive='left'
            )
            network.set_snapshots(snapshots)
            snapshots_length = len(snapshots)
        
        # 버스 추가 - 실제 데이터만 사용
        if 'buses' in input_data:
            print("\n=== 버스 추가 시작 ===")
            for _, bus in input_data['buses'].iterrows():
                bus_name = str(bus['name'])
                raw_carrier = str(bus['carrier'])
                token = _standardize_bus_token_by_carrier(raw_carrier)
                carrier_map = {'EL': 'electricity', 'H': 'heat', 'H2': 'hydrogen', 'LNG': 'gas'}
                carrier = carrier_map.get(token, raw_carrier)
                
                v_nom_val = float(bus['v_nom']) if pd.notna(bus['v_nom']) else 345.0
                network.add("Bus",
                          name=bus_name,
                          v_nom=v_nom_val,
                          carrier=carrier)
                print(f"버스 추가됨: {bus_name} (carrier: {carrier}, v_nom: {v_nom_val})")
        
        # 재생에너지 패턴 준비
        renewable_patterns = {}
        if 'renewable_patterns' in input_data:
            patterns_df = input_data['renewable_patterns']
            print(f"재생에너지 패턴 원본 데이터 확인:")
            print(f"컬럼: {patterns_df.columns.tolist()}")
            print(f"첫 10행:")
            print(patterns_df.head(10))
            
            # 실제 데이터가 있는 행 찾기 (4행부터 시작)
            if len(patterns_df) > 4:
                # 헤더 행 찾기 (3행: 시간, 태양광(PV), 풍력(WT))
                header_row = patterns_df.iloc[3]
                print(f"헤더 행: {header_row.tolist()}")
                
                # 데이터 행들 (4행부터)
                data_rows = patterns_df.iloc[4:].copy()
                
                # 컬럼명 재설정
                if len(header_row) >= 3:
                    data_rows.columns = ['시간', 'PV', 'WT']
                    
                    # PV 패턴 처리
                    if 'PV' in data_rows.columns:
                        pv_values = pd.to_numeric(data_rows['PV'], errors='coerce').dropna().values
                        if len(pv_values) > 0:
                            pv_pattern = normalize_pattern(pv_values)
                            pv_pattern = adjust_pattern_length(pv_pattern, snapshots_length)
                            renewable_patterns['PV_pattern'] = pv_pattern
                            print(f"PV 패턴 준비됨 - 길이: {len(pv_pattern)}, 최대값: {np.max(pv_pattern):.3f}, 최소값: {np.min(pv_pattern):.3f}")
                    
                    # WT 패턴 처리
                    if 'WT' in data_rows.columns:
                        wt_values = pd.to_numeric(data_rows['WT'], errors='coerce').dropna().values
                        if len(wt_values) > 0:
                            wt_pattern = normalize_pattern(wt_values)
                            wt_pattern = adjust_pattern_length(wt_pattern, snapshots_length)
                            renewable_patterns['WT_pattern'] = wt_pattern
                            print(f"WT 패턴 준비됨 - 길이: {len(wt_pattern)}, 최대값: {np.max(wt_pattern):.3f}, 최소값: {np.min(wt_pattern):.3f}")
            
            if not renewable_patterns:
                print("⚠️ 재생에너지 패턴을 찾을 수 없습니다. 기본값 1.0을 사용합니다.")
        
        # 발전기 추가
        if 'generators' in input_data:
            print("\n=== 발전기 추가 시작 ===")
            defined_buses = set(network.buses.index)
            bus_carriers_map = dict(zip(network.buses.index, network.buses.carrier))
            for _, gen in input_data['generators'].iterrows():
                gen_name = str(gen['name'])
                p_nom_value = float(gen['p_nom'])
                
                # p_nom이 0인 경우 건너뛰기
                if p_nom_value <= 0:
                    print(f"발전기 {gen_name} 건너뜀: p_nom이 0 이하 ({p_nom_value})")
                    continue
                
                raw_bus = str(gen['bus'])
                norm_bus = _normalize_bus_name(raw_bus, defined_buses, True, bus_carriers_map)
                if norm_bus != raw_bus:
                    print(f"발전기 {gen_name} 버스명 정규화: {raw_bus} → {norm_bus}")
                
                params = {
                    'name': gen_name,
                    'bus': norm_bus,
                    'p_nom': p_nom_value,
                    'carrier': str(gen['carrier'])
                }
                
                # 나머지 파라미터 추가
                optional_params = ['p_nom_extendable', 'marginal_cost', 'capital_cost', 'efficiency', 'p_min_pu', 'p_nom_min', 'p_nom_max']
                for param in optional_params:
                    if param in gen and pd.notna(gen[param]):
                        if param == 'p_nom_extendable':
                            params[param] = _to_bool(gen[param])
                        else:
                            params[param] = float(gen[param])
                
                network.add("Generator", **params)
                print(f"발전기 추가됨: {gen_name} (p_nom: {p_nom_value}, carrier: {params['carrier']})")
        
        # 발전기 추가 후 재생에너지 패턴 적용
        print("\n=== 재생에너지 패턴 적용 시작 ===")
        pv_applied_count = 0
        wt_applied_count = 0
        
        for gen_name in network.generators.index:
            pattern_applied = False
            
            # PV 패턴 적용 (더 넓은 범위로 매칭)
            if ('PV' in gen_name or 'Solar' in gen_name or 'solar' in gen_name) and 'PV_pattern' in renewable_patterns:
                # 패턴 길이 확인 및 조정
                pattern = renewable_patterns['PV_pattern']
                if len(pattern) != len(network.snapshots):
                    print(f"⚠️ PV 패턴 길이 불일치: {len(pattern)} vs {len(network.snapshots)}")
                    pattern = adjust_pattern_length(pattern, len(network.snapshots))
                
                network.generators_t.p_max_pu[gen_name] = pattern
                print(f"{gen_name}에 PV 패턴 적용됨 (길이: {len(pattern)}, 평균: {np.mean(pattern):.3f}, 최대: {np.max(pattern):.3f})")
                pv_applied_count += 1
                pattern_applied = True
            
            # WT 패턴 적용 (더 넓은 범위로 매칭)
            if ('WT' in gen_name or 'Wind' in gen_name or 'wind' in gen_name) and 'WT_pattern' in renewable_patterns:
                # 패턴 길이 확인 및 조정
                pattern = renewable_patterns['WT_pattern']
                if len(pattern) != len(network.snapshots):
                    print(f"⚠️ WT 패턴 길이 불일치: {len(pattern)} vs {len(network.snapshots)}")
                    pattern = adjust_pattern_length(pattern, len(network.snapshots))
                
                network.generators_t.p_max_pu[gen_name] = pattern
                print(f"{gen_name}에 WT 패턴 적용됨 (길이: {len(pattern)}, 평균: {np.mean(pattern):.3f}, 최대: {np.max(pattern):.3f})")
                wt_applied_count += 1
                pattern_applied = True
            
            # 패턴이 적용되지 않은 재생에너지 발전기 확인
            if not pattern_applied and any(keyword in gen_name.lower() for keyword in ['pv', 'solar', 'wind', 'wt']):
                print(f"⚠️ 재생에너지 발전기 {gen_name}에 패턴이 적용되지 않음")
        
        print(f"\n재생에너지 패턴 적용 완료:")
        print(f"- PV 패턴 적용: {pv_applied_count}개 발전기")
        print(f"- WT 패턴 적용: {wt_applied_count}개 발전기")
        
        # 적용된 패턴의 통계 정보 출력
        if 'PV_pattern' in renewable_patterns:
            pv_pattern = renewable_patterns['PV_pattern']
            print(f"- PV 패턴 통계: 최대 {np.max(pv_pattern):.3f}, 최소 {np.min(pv_pattern):.3f}, 평균 {np.mean(pv_pattern):.3f}")
        
        if 'WT_pattern' in renewable_patterns:
            wt_pattern = renewable_patterns['WT_pattern']
            print(f"- WT 패턴 통계: 최대 {np.max(wt_pattern):.3f}, 최소 {np.min(wt_pattern):.3f}, 평균 {np.mean(wt_pattern):.3f}")
        
        # 패턴 적용 후 검증
        print("\n=== 패턴 적용 검증 ===")
        for gen_name in network.generators.index:
            if any(keyword in gen_name.lower() for keyword in ['pv', 'solar', 'wind', 'wt']):
                if gen_name in network.generators_t.p_max_pu.columns:
                    pattern_values = network.generators_t.p_max_pu[gen_name]
                    print(f"{gen_name}: 패턴 적용됨 - 평균 {np.mean(pattern_values):.3f}, 변동성 {np.std(pattern_values):.3f}")
                else:
                    print(f"⚠️ {gen_name}: 패턴 적용되지 않음 - 기본값 1.0 사용")
        
        # 발전기 효율 및 최대/최소 출력비율(p_max_pu/p_min_pu) 반영
        try:
            gdf = input_data.get('generators', pd.DataFrame())
            if not gdf.empty:
                cols = list(gdf.columns)
                for _, row in gdf.iterrows():
                    gname = str(row.get('name', '')).strip()
                    if not gname or gname not in network.generators.index:
                        continue
                    # 효율
                    eff = pd.to_numeric(row.get('efficiency'), errors='coerce')
                    if pd.isna(eff) and '효율' in cols:
                        eff = pd.to_numeric(row.get('효율'), errors='coerce')
                    # 최대/최소 출력비율
                    max_ratio = pd.to_numeric(row.get('p_max_pu'), errors='coerce')
                    if pd.isna(max_ratio) and '최대출력비율' in cols:
                        max_ratio = pd.to_numeric(row.get('최대출력비율'), errors='coerce')
                    min_ratio = pd.to_numeric(row.get('p_min_pu'), errors='coerce')
                    if pd.isna(min_ratio) and '최소출력비율' in cols:
                        min_ratio = pd.to_numeric(row.get('최소출력비율'), errors='coerce')
                    # 상한 배율 = (기본 1.0) × max_ratio × eff(0~1)
                    cap_mul = 1.0
                    if pd.notna(max_ratio):
                        cap_mul = float(np.clip(max_ratio, 0.0, 1.0))
                    if pd.notna(eff) and 0.0 < float(eff) <= 1.0:
                        cap_mul *= float(eff)
                    # p_max_pu 적용(이미 패턴이 있으면 곱)
                    if gname in network.generators_t.p_max_pu.columns:
                        network.generators_t.p_max_pu[gname] = network.generators_t.p_max_pu[gname] * cap_mul
                    else:
                        network.generators_t.p_max_pu[gname] = pd.Series(cap_mul, index=network.snapshots)
                    # p_min_pu 적용(있을 때만)
                    if pd.notna(min_ratio):
                        min_val = float(np.clip(min_ratio, 0.0, 1.0))
                        network.generators_t.p_min_pu[gname] = pd.Series(min_val, index=network.snapshots)
                    try:
                        msg = f"발전기 출력제한 적용: {gname} (max×eff={cap_mul:.3f}"
                        if pd.notna(min_ratio):
                            msg += f", min={float(np.clip(min_ratio, 0.0, 1.0)):.3f}"
                        msg += ")"
                        print(msg)
                    except Exception:
                        pass
            # p_min_pu <= p_max_pu 보정
            try:
                for gname in network.generators.index:
                    if gname in network.generators_t.p_max_pu.columns:
                        pmax = network.generators_t.p_max_pu[gname].fillna(0.0)
                    else:
                        pmax = pd.Series(1.0, index=network.snapshots)
                    if hasattr(network.generators_t, 'p_min_pu') and (gname in network.generators_t.p_min_pu.columns):
                        pmin = network.generators_t.p_min_pu[gname].fillna(0.0)
                        network.generators_t.p_min_pu[gname] = pd.concat([pmin, pmax], axis=1).min(axis=1).clip(lower=0.0)
            except Exception as _e_clamp:
                print(f"p_min_pu 보정 경고: {_e_clamp}")
        except Exception as _e_der:
            print(f"발전기 효율/출력비율 적용 경고: {_e_der}")
        
        # 발전기 백업 보강: 각 버스에 최소 하나의 가용 발전기 보장
        try:
            print("\n=== 버스별 기본 발전기 보강 확인 ===")
            buses_with_gen = set(network.generators.bus.unique()) if not network.generators.empty else set()
            for bus in network.buses.index:
                # LNG 버스에는 전력 보강발전기 추가하지 않음(전력용)
                if bus.endswith('_LNG'):
                    # 대신 연료(가스) 공급원이 없다면 가스 버스에 연료공급 소스를 추가
                    try:
                        bus_carrier_chk = str(network.buses.at[bus, 'carrier']).lower()
                    except Exception:
                        bus_carrier_chk = ''
                    if bus_carrier_chk == 'gas':
                        fuel_name = f"{bus}_Fuel_Supply"
                        if fuel_name not in network.generators.index:
                            network.add("Generator",
                                       name=fuel_name,
                                       bus=bus,
                                       p_nom=0.0,
                                       p_nom_extendable=True,
                                       marginal_cost=0.0,
                                       carrier='gas')
                            network.generators_t.p_max_pu[fuel_name] = pd.Series(1.0, index=network.snapshots)
                            print(f"가스 연료공급 추가: {fuel_name} (버스 {bus})")
                    continue
                # 버스 캐리어 확인
                try:
                    bus_carrier = str(network.buses.at[bus, 'carrier']).lower()
                except Exception:
                    bus_carrier = ''
                # 열 버스: 항상 확장가능 고비용 보강발전기 추가(부족분만 보완)
                if bus_carrier == 'heat':
                    # 해당 열 버스에 실제 열부하가 있는 경우에만 보강발전기 후보 추가 (네트워크 시계열 기준)
                    has_heat_load = False
                    try:
                        total = 0.0
                        # 1) 네트워크 시계열 기준
                        if (not network.loads.empty) and (hasattr(network.loads_t, 'p_set') and not network.loads_t.p_set.empty):
                            load_names = network.loads.index[network.loads.bus.astype(str) == bus]
                            for nm in load_names:
                                if nm in network.loads_t.p_set.columns:
                                    total += float(np.nansum(network.loads_t.p_set[nm].values))
                        # 2) 폴백: 입력 데이터 기준
                        if total <= 0.0:
                            loads_df = input_data.get('loads', pd.DataFrame())
                            if not loads_df.empty and 'bus' in loads_df.columns and 'p_set' in loads_df.columns:
                                sub = loads_df[loads_df['bus'].astype(str) == bus]
                                if not sub.empty:
                                    total = float(pd.to_numeric(sub['p_set'], errors='coerce').fillna(0).sum())
                        has_heat_load = total > 0.0
                    except Exception:
                        has_heat_load = True
                    if not has_heat_load:
                        continue
                    fallback_name = f"{bus}_Fallback_Gen"
                    if fallback_name not in network.generators.index:
                        network.add("Generator",
                                   name=fallback_name,
                                   bus=bus,
                                   p_nom=0.0,
                                   p_nom_extendable=True,
                                   capital_cost=1e7,
                                   marginal_cost=1e6,
                                   carrier='heat')
                        network.generators_t.p_max_pu[fallback_name] = pd.Series(1.0, index=network.snapshots)
                        print(f"열 보강 발전기 추가(확장가능/매우고비용): {fallback_name} (버스 {bus})")
                    continue
                # 전력 버스: LNG 보완발전기(확장가능/매우고비용) 추가 — 실제 전력부하가 있을 때만
                if bus_carrier == 'electricity' or bus_carrier == 'AC':
                    has_el_load = False
                    try:
                        total = 0.0
                        # 1) 네트워크 시계열 기준
                        if (not network.loads.empty) and (hasattr(network.loads_t, 'p_set') and not network.loads_t.p_set.empty):
                            load_names = network.loads.index[network.loads.bus.astype(str) == bus]
                            for nm in load_names:
                                if nm in network.loads_t.p_set.columns:
                                    total += float(np.nansum(network.loads_t.p_set[nm].values))
                        # 2) 폴백: 입력 데이터 기준
                        if total <= 0.0:
                            loads_df = input_data.get('loads', pd.DataFrame())
                            if not loads_df.empty and 'bus' in loads_df.columns and 'p_set' in loads_df.columns:
                                sub = loads_df[loads_df['bus'].astype(str) == bus]
                                if not sub.empty:
                                    total = float(pd.to_numeric(sub['p_set'], errors='coerce').fillna(0).sum())
                        has_el_load = total > 0.0
                    except Exception:
                        has_el_load = True
                    if not has_el_load:
                        continue
                    fallback_name = f"{bus}_LNG_Fallback_Gen"
                    if fallback_name not in network.generators.index:
                        network.add("Generator",
                                   name=fallback_name,
                                   bus=bus,
                                   p_nom=0.0,
                                   p_nom_extendable=True,
                                   capital_cost=1e7,
                                   marginal_cost=1e6,
                                   carrier='gas')
                        network.generators_t.p_max_pu[fallback_name] = pd.Series(1.0, index=network.snapshots)
                        print(f"전력 보완 발전기 추가(확장가능/매우고비용): {fallback_name} (버스 {bus})")
                    # 전력 슬랙 발전기(무한 확장/초고비용) 추가 - infeasible 방지를 위해 기본 활성화
                    try:
                        # infeasible 방지를 위해 기본적으로 활성화 (DISABLE_POWER_SLACK=1로 비활성화 가능)
                        disable_slack = os.environ.get('DISABLE_POWER_SLACK', '0')
                        print(f"DEBUG: DISABLE_POWER_SLACK={disable_slack}, 버스 {bus}")
                        if disable_slack != '1':
                            slack_name = f"{bus}_Slack_Gen"
                            if slack_name not in network.generators.index:
                                slack_cost = float(os.environ.get('SLACK_GEN_COST', '1e9'))
                                network.add("Generator",
                                           name=slack_name,
                                           bus=bus,
                                           p_nom=0.0,
                                           p_nom_extendable=True,
                                           p_nom_min=0.0,
                                           p_nom_max=1e6,  # 매우 큰 확장 한계
                                           capital_cost=0.0,
                                           marginal_cost=slack_cost,
                                           carrier='AC')
                                network.generators_t.p_max_pu[slack_name] = pd.Series(1.0, index=network.snapshots)
                                print(f"전력 슬랙 발전기 추가(infeasible 방지): {slack_name} (버스 {bus}, mcost={slack_cost})")
                    except Exception as _e_sl:
                        print(f"슬랙 발전기 추가 경고: {_e_sl}")
                # 수소 버스: 수소 백업기(필요시)
                if bus_carrier == 'hydrogen':
                    total_load = 0.0
                    try:
                        if (not network.loads.empty) and hasattr(network.loads_t, 'p_set') and (network.loads_t.p_set is not None) and (not network.loads_t.p_set.empty):
                            load_names = network.loads.index[network.loads.bus.astype(str) == bus]
                            for nm in load_names:
                                if nm in network.loads_t.p_set.columns:
                                    total_load += float(np.nansum(network.loads_t.p_set[nm].values))
                        if total_load <= 0.0 and (not network.loads.empty) and hasattr(network.loads_t, 'p') and (network.loads_t.p is not None) and (not network.loads_t.p.empty):
                            load_names = network.loads.index[network.loads.bus.astype(str) == bus]
                            for nm in load_names:
                                if nm in network.loads_t.p.columns:
                                    total_load += float(np.nansum(network.loads_t.p[nm].values))
                        if total_load <= 0.0:
                            loads_df = input_data.get('loads', pd.DataFrame())
                            if not loads_df.empty and 'bus' in loads_df.columns and 'p_set' in loads_df.columns:
                                sub = loads_df[loads_df['bus'].astype(str) == bus]
                                if not sub.empty:
                                    total_load = float(pd.to_numeric(sub['p_set'], errors='coerce').fillna(0).sum())
                    except Exception:
                        pass
                    if total_load > 0.0:
                        fallback_name = f"{bus}_H2_Fallback_Gen"
                        if fallback_name not in network.generators.index:
                            network.add("Generator",
                                       name=fallback_name,
                                       bus=bus,
                                       p_nom=0.0,
                                       p_nom_extendable=True,
                                       capital_cost=1e7,
                                       marginal_cost=1e6,
                                       carrier='hydrogen')
                            network.generators_t.p_max_pu[fallback_name] = pd.Series(1.0, index=network.snapshots)
        except Exception as e:
            print(f"기본 발전기 보강 중 오류: {str(e)}")

        # 모든 발전기에 p_max_pu 기본값(1.0) 보장
        try:
            missing_cols = [g for g in network.generators.index if g not in network.generators_t.p_max_pu.columns]
            if missing_cols:
                for g in missing_cols:
                    network.generators_t.p_max_pu[g] = pd.Series(1.0, index=network.snapshots)
                print(f"p_max_pu 기본 적용: {len(missing_cols)}개 발전기에 1.0 설정")
        except Exception as e:
            print(f"p_max_pu 기본값 설정 중 오류: {str(e)}")
        
        # 부하 추가
        if 'loads' in input_data:
            print("\n=== 부하 추가 시작 ===")
            for _, load in input_data['loads'].iterrows():
                name = str(load['name'])
                bus_name = str(load['bus'])
                p_set = float(load['p_set'])
                
                # 부하 패턴 적용
                if 'load_patterns' in input_data:
                    # 부하 이름에서 지역과 타입 추출
                    region = name.split('_')[0] if '_' in name else None
                    dtype = 'EL' if '_Demand_EL' in name else ('H2' if '_Demand_H2' in name else ('H' if '_Demand_H' in name else None))
                    pattern = _get_load_pattern(input_data, region, dtype, len(snapshots)) if region and dtype else None

                    if pattern is not None:
                        # 총수요 × 8760 × 패턴(스케일 없이 그대로 적용)
                        p_set = float(p_set) * 8760.0 * pattern
                        print(f"부하 {name}에 패턴 적용됨 (총수요×8760×패턴)")
                    else:
                        cols = list(input_data['load_patterns'].columns) if 'load_patterns' in input_data else []
                        print(f"부하 {name}: 패턴 미발견 (region={region}, type={dtype}), 사용 가능 컬럼: {cols[:10]}{'...' if len(cols)>10 else ''}")
                        # 패턴이 없는 경우 일정한 부하
                        p_set = np.full(len(snapshots), p_set)
                        print(f"부하 {name}에 일정한 부하 적용됨")
                else:
                    p_set = np.full(len(snapshots), p_set)
                
                network.add("Load",
                          name=name,
                          bus=bus_name)
                # 시간별 p_set은 명시적으로 loads_t에 설정 (스냅샷 인덱스 정렬)
                if isinstance(p_set, pd.Series):
                    series = p_set.reindex(network.snapshots)
                elif isinstance(p_set, (np.ndarray, list)):
                    series = pd.Series(p_set, index=network.snapshots)
                else:
                    series = pd.Series(np.full(len(network.snapshots), float(p_set)), index=network.snapshots)
                network.loads_t.p_set[name] = series
                print(f"부하 추가됨: {name} (버스: {bus_name})")
        
        # 시나리오 수요 스케일링 비활성화 (지역별 시트 원본 데이터 사용)
        # _apply_scenario_demand_scaling(network, input_data)
        
        # 스케일링 이후 사후 백업 발전기 보강(전력/열 버스 대상)
        try:
            added_backup = 0
            for bus in network.buses.index:
                try:
                    bus_carrier_post = str(network.buses.at[bus, 'carrier']).lower()
                except Exception:
                    bus_carrier_post = ''
                # 버스별 부하 총량 계산(p_set 우선, 없으면 p)
                total_load = 0.0
                try:
                    if (not network.loads.empty) and hasattr(network.loads_t, 'p_set') and (network.loads_t.p_set is not None) and (not network.loads_t.p_set.empty):
                        load_names = network.loads.index[network.loads.bus.astype(str) == bus]
                        for nm in load_names:
                            if nm in network.loads_t.p_set.columns:
                                total_load += float(np.nansum(network.loads_t.p_set[nm].values))
                    if total_load <= 0.0 and (not network.loads.empty) and hasattr(network.loads_t, 'p') and (network.loads_t.p is not None) and (not network.loads_t.p.empty):
                        load_names = network.loads.index[network.loads.bus.astype(str) == bus]
                        for nm in load_names:
                            if nm in network.loads_t.p.columns:
                                total_load += float(np.nansum(network.loads_t.p[nm].values))
                    # 최후 폴백: 입력 데이터 기준
                    if total_load <= 0.0:
                        loads_df2 = input_data.get('loads', pd.DataFrame())
                        if not loads_df2.empty and 'bus' in loads_df2.columns and 'p_set' in loads_df2.columns:
                            sub2 = loads_df2[loads_df2['bus'].astype(str) == bus]
                            if not sub2.empty:
                                total_load = float(pd.to_numeric(sub2['p_set'], errors='coerce').fillna(0).sum())
                except Exception:
                    pass
                # 열 버스: 열 백업기
                if bus_carrier_post == 'heat' and total_load > 0.0:
                    fallback_name = f"{bus}_Fallback_Gen"
                    if fallback_name not in network.generators.index:
                        network.add("Generator",
                                   name=fallback_name,
                                   bus=bus,
                                   p_nom=0.0,
                                   p_nom_extendable=True,
                                   capital_cost=1e7,
                                   marginal_cost=1e6,
                                   carrier='heat')
                        network.generators_t.p_max_pu[fallback_name] = pd.Series(1.0, index=network.snapshots)
                        added_backup += 1
                # 전력 버스: LNG 백업기
                if bus_carrier_post == 'electricity' and total_load > 0.0:
                    fallback_name = f"{bus}_LNG_Fallback_Gen"
                    if fallback_name not in network.generators.index:
                        network.add("Generator",
                                   name=fallback_name,
                                   bus=bus,
                                   p_nom=0.0,
                                   p_nom_extendable=True,
                                   capital_cost=1e7,
                                   marginal_cost=1e6,
                                   carrier='gas')
                        network.generators_t.p_max_pu[fallback_name] = pd.Series(1.0, index=network.snapshots)
                        added_backup += 1
                # 수소 버스: 수소 백업기(필요시)
                if bus_carrier_post == 'hydrogen' and total_load > 0.0:
                    fallback_name = f"{bus}_H2_Fallback_Gen"
                    if fallback_name not in network.generators.index:
                        network.add("Generator",
                                   name=fallback_name,
                                   bus=bus,
                                   p_nom=0.0,
                                   p_nom_extendable=True,
                                   capital_cost=1e7,
                                   marginal_cost=1e6,
                                   carrier='hydrogen')
                        network.generators_t.p_max_pu[fallback_name] = pd.Series(1.0, index=network.snapshots)
                        added_backup += 1
            if added_backup > 0:
                print(f"스케일링 이후 백업 발전기 보강: {added_backup}개 추가")
        except Exception as _e_post:
            print(f"사후 백업 발전기 보강 경고: {_e_post}")
        
        # Links 추가
        if 'links' in input_data:
            print("\n=== Links 추가 시작 ===")
            links_df = input_data['links']
            print(f"Links 시트 컬럼: {list(links_df.columns)}")
            
            # CHP 제외 모드 제거됨
            
            for idx, link in links_df.iterrows():
                link_name = str(link['name'])
                print(f"\nLink 처리 중: {link_name}")
                
                # CHP 판단: bus2가 있거나 이름에 'CHP' 포함
                # CHP 필터링 제거
                
                # 후보 컬럼에서 값 읽기 유틸
                def _read_first_available(col_candidates):
                    for c in col_candidates:
                        if (c in links_df.columns) and pd.notna(link.get(c)):
                            return link.get(c)
                    return None
                
                # 버스 연결 읽기(한글/영문 모두 지원)
                bus0_raw = _read_first_available(['bus0', '시작버스', 'from', '출발', '시작'])
                bus1_raw = _read_first_available(['bus1', '종료버스', 'to', '도착', '끝'])
                bus2_raw = _read_first_available(['bus2', '버스2'])
                bus3_raw = _read_first_available(['bus3', '버스3'])
                
                bus0_name = str(bus0_raw) if bus0_raw is not None else None
                bus1_name = str(bus1_raw) if bus1_raw is not None else None
                bus2_name = str(bus2_raw) if bus2_raw is not None else None
                bus3_name = str(bus3_raw) if bus3_raw is not None else None
                
                # 정의된 버스들만 확인
                defined_buses = set(network.buses.index)
                bus_carriers = dict(zip(network.buses.index, network.buses.carrier))
                # 이름 정규화(예: BSN_EL ↔ BSN_BSN_EL)
                # 주의: 연료 버스(LNG/gas)는 EL/H/H2로 강제 치환하지 않도록 전력 선호(prefer_electric)를 끕니다.
                if bus0_name:
                    prefer = False if (('lng' in bus0_name.lower()) or ('gas' in bus0_name.lower())) else True
                    bus0_name = _normalize_bus_name(bus0_name, defined_buses, prefer, bus_carriers)
                if bus1_name:
                    bus1_name = _normalize_bus_name(bus1_name, defined_buses, True, bus_carriers)
                if bus2_name:
                    prefer2 = False if (('lng' in bus2_name.lower()) or ('gas' in bus2_name.lower())) else True
                    bus2_name = _normalize_bus_name(bus2_name, defined_buses, prefer2, bus_carriers)
                if bus3_name:
                    prefer3 = False if (('lng' in bus3_name.lower()) or ('gas' in bus3_name.lower())) else True
                    bus3_name = _normalize_bus_name(bus3_name, defined_buses, prefer3, bus_carriers)

                # 링크 유형에 따른 버스 자동 정렬/보정
                try:
                    lname = link_name.lower()
                    is_chp_link = False
                    def _is_el(b):
                        return bool(b) and (b.endswith('_EL') or str(bus_carriers.get(b, '')).lower() == 'electricity')
                    def _is_h(b):
                        return bool(b) and (b.endswith('_H') or str(bus_carriers.get(b, '')).lower() == 'heat')
                    def _is_h2(b):
                        return bool(b) and (b.endswith('_H2') or str(bus_carriers.get(b, '')).lower() == 'hydrogen')
                    def _is_gas(b):
                        return bool(b) and (b.endswith('_LNG') or 'gas' in str(bus_carriers.get(b, '')).lower())

                    # CHP: bus0=연료(LNG/gas), bus1=전기, bus2=열
                    cand = [bus0_name, bus1_name, bus2_name]
                    fuel = next((b for b in cand if _is_gas(b)), None)
                    el   = next((b for b in cand if _is_el(b)), None)
                    heat = next((b for b in cand if _is_h(b)), None)
                    # 포트 패턴이 보이면 이름 여부와 무관하게 항상 CHP로 정렬
                    if (fuel and el and heat) or ('chp' in lname):
                        if fuel and el and heat:
                            if (bus0_name, bus1_name, bus2_name) != (fuel, el, heat):
                                note = "이름 기반" if ('chp' in lname) else "포트 기반"
                                print(f"CHP 버스 자동정렬({note}): {bus0_name},{bus1_name},{bus2_name} -> {fuel},{el},{heat}")
                            bus0_name, bus1_name, bus2_name = fuel, el, heat
                        is_chp_link = True

                    # Electrolyser: bus0=전기, bus1=수소, bus2 미사용
                    elif ('electrolyser' in lname) or ('electrolyzer' in lname):
                        if not _is_el(bus0_name) and _is_el(bus1_name):
                            print(f"Electrolyser 버스 스왑: {bus0_name}<->{bus1_name}")
                            bus0_name, bus1_name = bus1_name, bus0_name
                        bus2_name = None

                    # Heat Pump: bus0=전기, bus1=열, bus2 미사용
                    elif ('heatpump' in lname) or ('heat_pump' in lname) or (lname.startswith('hp') or ' hp ' in lname):
                        if not _is_el(bus0_name) and _is_el(bus1_name):
                            print(f"HeatPump 버스 스왑: {bus0_name}<->{bus1_name}")
                            bus0_name, bus1_name = bus1_name, bus0_name
                        bus2_name = None
                except Exception:
                    pass
                
                # 유효한 버스 연결 확인 (bus0/bus1는 필수)
                if bus0_name and bus1_name and bus0_name in defined_buses and bus1_name in defined_buses:
                    # 효율(한글/영문 매핑):
                    # - efficiency: bus1로의 효율(우선순위: 'efficiency' → 'efficiency0' → '효율0')
                    # - efficiency2: bus2로의 효율(우선순위: 'efficiency2' → 'efficiency1' → 'efficiency_2' → '효율2' → '효율1')
                    # - efficiency3: bus3로의 효율(우선순위: 'efficiency3' → 'efficiency_3' → '효율3')
                    eff1_val = _read_first_available(['efficiency', 'efficiency0', '효율0'])
                    eff2_val = _read_first_available(['efficiency2', 'efficiency1', 'efficiency_2', '효율2', '효율1'])
                    eff3_val = _read_first_available(['efficiency3', 'efficiency_3', '효율3'])
                    
                    # Link 정격 용량(p_nom) 읽기: 다양한 대체 컬럼 지원
                    pnom_raw = _read_first_available(['p_nom', '용량', 'capacity', '정격', '전력용량', '전력_용량'])
                    try:
                        pnom_val = float(pd.to_numeric(pnom_raw, errors='coerce')) if pnom_raw is not None else float('nan')
                    except Exception:
                        pnom_val = float('nan')
                    if (pnom_raw is None) or (pd.isna(pnom_val)):
                        # 시트에 p_nom 컬럼이 있는지 확인
                        if 'p_nom' in links_df.columns:
                            sheet_pnom = link.get('p_nom')
                            if pd.notna(sheet_pnom):
                                pnom_val = float(sheet_pnom)
                            else:
                                # NaN인 경우, HP는 0으로 기본 설정, 다른 링크는 100
                                if 'hp' in link_name.lower():
                                    pnom_val = 0.0  # HP는 0으로 기본 설정
                                    print(f"HP 링크 {link_name}: p_nom이 NaN이므로 0으로 설정")
                                else:
                                    pnom_val = 100.0
                        else:
                            pnom_val = 100.0  # 진짜로 컬럼이 없을 때만 기본값 사용
                    
                    params = {
                        'name': link_name,
                        'bus0': bus0_name,
                        'bus1': bus1_name,
                        'p_nom': pnom_val,
                        'efficiency': float(eff1_val) if pd.notna(eff1_val) else (0.5 if ('electrolyser' in link_name.lower() or 'electrolyzer' in link_name.lower()) else 0.9)
                    }

                    # 열 공급 우선순위 유도: CHP < HP < Fallback
                    try:
                        default_mcost = None
                        lname2 = link_name.lower()
                        # 시트에 명시된 marginal_cost가 없을 때만 기본값 적용
                        has_sheet_mcost = ('marginal_cost' in links_df.columns and pd.notna(link.get('marginal_cost')))
                        if not has_sheet_mcost:
                            if is_chp_link:
                                default_mcost = 1.0  # 최우선 사용
                            elif ('heatpump' in lname2) or ('heat_pump' in lname2) or (lname2.startswith('hp') or ' hp ' in lname2):
                                default_mcost = 1e5  # Fallback보다 낮고 CHP보다 높게
                        if default_mcost is not None:
                            params['marginal_cost'] = float(default_mcost)
                        elif has_sheet_mcost:
                            params['marginal_cost'] = float(link.get('marginal_cost'))
                    except Exception:
                        pass
                    
                    # 선택적 추가 출력 버스와 효율
                    if bus2_name and bus2_name in defined_buses:
                        params['bus2'] = bus2_name
                        if pd.notna(eff2_val):
                            params['efficiency2'] = float(eff2_val)
                        else:
                            # CHP로 추정되나 eff2가 비어 있으면 기본 0.4 적용
                            # NaN 방지: CHP 포함 모든 경우, 명시 없으면 0.0으로 설정
                            params['efficiency2'] = 0.0
                    elif bus2_name:
                        print(f"경고: Link {link_name}의 bus2 '{bus2_name}'가 네트워크에 없어 무시합니다.")
                    
                    if bus3_name and bus3_name in defined_buses:
                        params['bus3'] = bus3_name
                        if pd.notna(eff3_val):
                            params['efficiency3'] = float(eff3_val)
                        else:
                            # NaN 방지
                            params['efficiency3'] = 0.0
                    elif bus3_name:
                        print(f"경고: Link {link_name}의 bus3 '{bus3_name}'가 네트워크에 없어 무시합니다.")
                    
                    # p_nom_extendable 및 용량 하한/상한 처리
                    if 'p_nom_extendable' in links_df.columns and pd.notna(link.get('p_nom_extendable')):
                        params['p_nom_extendable'] = _to_bool(link.get('p_nom_extendable'))
                    if 'p_nom_min' in links_df.columns and pd.notna(link.get('p_nom_min')):
                        params['p_nom_min'] = float(link.get('p_nom_min'))
                    if 'p_nom_max' in links_df.columns and pd.notna(link.get('p_nom_max')):
                        params['p_nom_max'] = float(link.get('p_nom_max'))
                    
                    try:
                        network.add("Link", **params)
                        print_msg = f"Link {link_name} 추가됨: {bus0_name} -> {bus1_name}"
                        if 'bus2' in params:
                            print_msg += f", bus2: {params['bus2']}"
                        if 'bus3' in params:
                            print_msg += f", bus3: {params['bus3']}"
                        print_msg += f" (p_nom: {params.get('p_nom', 'n/a')}, eff: {params.get('efficiency', 'n/a')}"
                        if 'efficiency2' in params:
                            print_msg += f", eff2: {params['efficiency2']}"
                        if 'efficiency3' in params:
                            print_msg += f", eff3: {params['efficiency3']}"
                        if is_chp_link:
                            print_msg += ", CHP=Y"
                        print_msg += ")"
                        print(print_msg)
                    except Exception as e:
                        print(f"Link {link_name} 추가 중 오류: {str(e)}")
                else:
                    print(f"Link {link_name} 건너뜀: 유효하지 않은 버스 연결 (bus0: {bus0_name}, bus1: {bus1_name})")
        
        # 저장장치 추가
        if 'stores' in input_data:
            print("\n=== Stores 추가 시작 ===")
            for _, store in input_data['stores'].iterrows():
                store_name = str(store['name'])
                bus_name = str(store['bus'])
                
                # 버스 존재 확인
                if bus_name in network.buses.index:
                    params = {
                        'name': store_name,
                        'bus': bus_name,
                        'carrier': str(store['carrier']),
                        'e_nom': float(store['e_nom']) if pd.notna(store['e_nom']) else 0,
                        'e_cyclic': _to_bool(store['e_cyclic']) if 'e_cyclic' in store and pd.notna(store['e_cyclic']) else True,
                        'efficiency_store': float(store['efficiency_store']) if 'efficiency_store' in store and pd.notna(store['efficiency_store']) else 0.9,
                        'efficiency_dispatch': float(store['efficiency_dispatch']) if 'efficiency_dispatch' in store and pd.notna(store['efficiency_dispatch']) else 0.9,
                        'standing_loss': float(store['standing_loss']) if 'standing_loss' in store and pd.notna(store['standing_loss']) else 0,
                        'e_initial': float(store['e_initial']) if 'e_initial' in store and pd.notna(store['e_initial']) else 0
                    }
                    if 'e_nom_extendable' in store and pd.notna(store['e_nom_extendable']):
                        params['e_nom_extendable'] = _to_bool(store['e_nom_extendable'])
                    
                    if 'e_nom_min' in store and pd.notna(store['e_nom_min']):
                        params['e_nom_min'] = float(store['e_nom_min'])
                    if 'e_nom_max' in store and pd.notna(store['e_nom_max']):
                        params['e_nom_max'] = float(store['e_nom_max'])
                    
                    network.add("Store", **params)
                    print(f"저장장치 {store_name} 추가됨 (버스: {bus_name})")
                else:
                    print(f"저장장치 {store_name} 건너뜀: 버스 '{bus_name}'가 존재하지 않음")
        
        # 선로 추가 (있는 경우)
        if 'lines' in input_data and not input_data['lines'].empty:
            print("\n=== 선로 추가 시작 ===")
            added_lines = 0
            skipped_lines = 0
            
            for _, line in input_data['lines'].iterrows():
                line_name = str(line['name'])
                bus0_name = str(line['bus0'])
                bus1_name = str(line['bus1'])
                
                # 버스 이름 정규화
                defined_buses = set(network.buses.index)
                bus_carriers = dict(zip(network.buses.index, network.buses.carrier))
                bus0_name_norm = _normalize_bus_name(bus0_name, defined_buses, True, bus_carriers)
                bus1_name_norm = _normalize_bus_name(bus1_name, defined_buses, True, bus_carriers)
                
                # 선로 이름에서 지역코드 강제 추출 후 전력버스로 강제 매핑 시도 (예: GND_JND → GND_EL, JND_EL)

                
                print(f"선로 {line_name} 시도: {bus0_name} -> {bus1_name} (정규화/강제: {bus0_name_norm} -> {bus1_name_norm})")
                print(f"  bus0 존재: {bus0_name_norm in network.buses.index}")
                print(f"  bus1 존재: {bus1_name_norm in network.buses.index}")
                
                if bus0_name_norm in network.buses.index and bus1_name_norm in network.buses.index:
                    # 선로 전압(type) 기반으로 각 끝단에 해당 전압 보조 버스를 확보하고 변압기(링크) 자동 삽입
                    def _bus_v(bus_name):
                        try:
                            return float(network.buses.at[bus_name, 'v_nom'])
                        except Exception:
                            return float('nan')

                    def _parse_voltage(val):
                        try:
                            if pd.isna(val):
                                return None
                            s = str(val)
                            nums = ''.join(ch for ch in s if ch.isdigit())
                            if nums:
                                return float(nums)
                        except Exception:
                            pass
                        return None

                    line_voltage = _parse_voltage(line['type']) if ('type' in line) else None

                    # 라인 파라미터 구성
                    params = {
                        'name': line_name,
                        'bus0': bus0_name_norm,
                        'bus1': bus1_name_norm,
                        's_nom': float(line['s_nom']) if ('s_nom' in line and pd.notna(line['s_nom'])) else 1000.0,
                        'x': float(line['x']) if ('x' in line and pd.notna(line['x'])) else 0.1,
                        'r': float(line['r']) if ('r' in line and pd.notna(line['r'])) else 0.01
                    }

                    # 추가 속성: type, length, num_parallel
                    if 'type' in line and pd.notna(line['type']):
                        tval = str(line['type']).strip()
                        try:
                            available_types = set(network.line_types.index)
                        except Exception:
                            available_types = set()
                        if tval in available_types:
                            params['type'] = tval
                    if 'length' in line and pd.notna(line['length']):
                        params['length'] = float(line['length'])
                    if 'num_parallel' in line and pd.notna(line['num_parallel']):
                        params['num_parallel'] = int(pd.to_numeric(line['num_parallel'], errors='coerce'))

                    # s_nom 확장 및 하한/상한 처리
                    if 's_nom_extendable' in line and pd.notna(line['s_nom_extendable']):
                        params['s_nom_extendable'] = bool(line['s_nom_extendable'])
                    if 's_nom_min' in line and pd.notna(line['s_nom_min']):
                        params['s_nom_min'] = float(line['s_nom_min'])
                    if 's_nom_max' in line and pd.notna(line['s_nom_max']):
                        params['s_nom_max'] = float(line['s_nom_max'])

                    # 라인 전압 기반 변압기 자동삽입(기본 비활성화). 활성화하려면 ENABLE_AUTO_TRANSFORMER=1 환경변수 설정
                    if (os.environ.get('ENABLE_AUTO_TRANSFORMER', '0') == '1') and (line_voltage is not None and not np.isnan(line_voltage)):
                        def ensure_voltage_bus(orig_bus, target_kv):
                            if orig_bus not in network.buses.index:
                                return orig_bus
                            if abs(_bus_v(orig_bus) - target_kv) < 1e-6:
                                return orig_bus
                            new_bus = f"{orig_bus}_{int(target_kv)}"
                            if new_bus not in network.buses.index:
                                network.add("Bus", name=new_bus, v_nom=target_kv, carrier='electricity')
                                # 변압기 등가: Link로 근사(효율 0.995)
                                tr_name = f"TR_{orig_bus}_to_{new_bus}"
                                if tr_name not in network.links.index:
                                    network.add("Link", name=tr_name, bus0=orig_bus, bus1=new_bus, p_nom=1e6, efficiency=0.995)
                                # 역방향 변압기 링크도 추가(양방향 전력 흐름 허용)
                                tr_rev = f"TR_{new_bus}_to_{orig_bus}"
                                if tr_rev not in network.links.index:
                                    network.add("Link", name=tr_rev, bus0=new_bus, bus1=orig_bus, p_nom=1e6, efficiency=0.995)
                            return new_bus
                        params['bus0'] = ensure_voltage_bus(bus0_name_norm, line_voltage)
                        params['bus1'] = ensure_voltage_bus(bus1_name_norm, line_voltage)

                    try:
                        network.add("Line", **params)
                        print(f"선로 {line_name} 추가됨: {params['bus0']} - {params['bus1']}")
                        added_lines += 1
                    except Exception as e:
                        print(f"선로 {line_name} 추가 실패: {str(e)}")
                        skipped_lines += 1
                else:
                    print(f"선로 {line_name} 건너뜀: 유효하지 않은 버스 연결")
                    skipped_lines += 1
                    if bus0_name not in network.buses.index:
                        print(f"  누락된 버스: {bus0_name}")
                    if bus1_name not in network.buses.index:
                        print(f"  누락된 버스: {bus1_name}")
            
            print(f"\n선로 추가 요약: 성공 {added_lines}개, 실패 {skipped_lines}개")
            print(f"네트워크에 추가된 총 선로 수: {len(network.lines)}")
            
            # 버스 목록 확인
            print(f"\n현재 네트워크 버스 수: {len(network.buses)}")
            print("첫 10개 버스:", list(network.buses.index[:10]))
            
            # CHP 제외 요약 제거됨
        else:
            print("\n선로 데이터가 없습니다.")
        
        # 링크 효율 NaN 일괄 보정: efficiency2/efficiency3 NaN → 0.0
        try:
            if not network.links.empty:
                if 'efficiency2' in network.links.columns:
                    network.links['efficiency2'] = pd.to_numeric(network.links['efficiency2'], errors='coerce').fillna(0.0)
                if 'efficiency3' in network.links.columns:
                    network.links['efficiency3'] = pd.to_numeric(network.links['efficiency3'], errors='coerce').fillna(0.0)
                print("링크 효율 NaN 보정 완료(efficiency2/3 → 0.0)")
        except Exception as _e_eff:
            print(f"링크 효율 보정 경고: {_e_eff}")
        
        # CO2 제약 추가 (한글/영문 헤더 모두 지원)
        if os.environ.get('DISABLE_CO2_LIMIT','0')=='1':
            print('CO2 제약 비활성화됨(DISABLE_CO2_LIMIT=1)')
        elif 'constraints' in input_data and not input_data['constraints'].empty:
            constraints_df = input_data['constraints'].copy()
            # 컬럼 표준화
            rename_map = {
                '이름': 'name', 'name': 'name',
                '상수': 'constant', 'constant': 'constant',
                '조건': 'sense', 'sense': 'sense',
                '유형': 'type', 'type': 'type',
                '대상속성': 'carrier_attribute', 'carrier_attribute': 'carrier_attribute'
            }
            std_cols = {}
            for c in constraints_df.columns:
                key = str(c).strip()
                std_cols[c] = rename_map.get(key, key)
            constraints_df.rename(columns=std_cols, inplace=True)

            co2_limit = pd.DataFrame()
            if 'name' in constraints_df.columns:
                co2_limit = constraints_df[constraints_df['name'].astype(str).str.strip() == 'CO2Limit']

            if not co2_limit.empty:
                const_col = 'constant' if 'constant' in co2_limit.columns else None
                if const_col:
                    limit_value = float(pd.to_numeric(co2_limit.iloc[0][const_col], errors='coerce'))
                    network.add("GlobalConstraint",
                                "CO2Limit",
                                sense="<=",
                                constant=limit_value)
                    print(f"CO2 제약이 {limit_value} tCO2로 설정되었습니다.")
                else:
                    print("경고: constraints 시트에 'constant' 컬럼이 없어 CO2Limit을 적용하지 못했습니다.")
            else:
                print("경고: CO2Limit 행이 없습니다.")
        else:
            print("경고: constraints 시트에 'name' 컬럼이 없어 전역 제약을 적용하지 못했습니다.")
        
        # 최종 안전장치: 여전히 수요 충족이 불가할 경우 초고비용 슬랙 발전기 추가
        try:
            failsafe_added = 0
            ensure_slack = os.environ.get('ENABLE_ALWAYS_SLACK', '1') == '1'
            for bus in network.buses.index:
                try:
                    bus_carrier = str(network.buses.at[bus, 'carrier']).lower()
                except Exception:
                    bus_carrier = ''
                if bus_carrier not in ['electricity', 'heat', 'hydrogen']:
                    continue
                # 해당 버스 부하 존재 여부
                has_load = False
                try:
                    if (not network.loads.empty) and hasattr(network.loads_t, 'p_set') and (network.loads_t.p_set is not None) and (not network.loads_t.p_set.empty):
                        names = network.loads.index[network.loads.bus.astype(str) == bus]
                        for nm in names:
                            if nm in network.loads_t.p_set.columns and float(np.nansum(network.loads_t.p_set[nm].values)) > 0.0:
                                has_load = True
                                break
                    if (not has_load) and (not network.loads.empty) and hasattr(network.loads_t, 'p') and (network.loads_t.p is not None) and (not network.loads_t.p.empty):
                        names = network.loads.index[network.loads.bus.astype(str) == bus]
                        for nm in names:
                            if nm in network.loads_t.p.columns and float(np.nansum(network.loads_t.p[nm].values)) > 0.0:
                                has_load = True
                                break
                except Exception:
                    pass
                if not has_load:
                    continue
                # 기존 발전기 유무 확인
                has_gen = False
                try:
                    gens = network.generators.index[network.generators.bus.astype(str) == bus]
                    if len(gens) > 0:
                        has_gen = True
                except Exception:
                    pass
                # ensure_slack가 활성화되면 기존 발전기가 있어도 슬랙 추가
                if (not has_gen) or ensure_slack:
                    slack_name = f"{bus}_Slack_Failsafe"
                    if slack_name not in network.generators.index:
                        mcost = float(os.environ.get('SLACK_GEN_COST', '1e9'))
                        # 전력/열/수소 중 해당 캐리어로 무배출 초고비용 슬랙
                        carrier_val = bus_carrier if bus_carrier in ['electricity','heat','hydrogen'] else 'electricity'
                        network.add("Generator",
                                   name=slack_name,
                                   bus=bus,
                                   p_nom=0.0,
                                   p_nom_extendable=True,
                                   capital_cost=0.0,
                                   marginal_cost=mcost,
                                   carrier=carrier_val)
                        network.generators_t.p_max_pu[slack_name] = pd.Series(1.0, index=network.snapshots)
                        failsafe_added += 1
            if failsafe_added > 0:
                print(f"최후수단 슬랙 발전기 추가: {failsafe_added}개")
        except Exception as _e_fs:
            print(f"최후수단 슬랙 발전기 추가 경고: {_e_fs}")
        
        # 경계값(최소/최대) 정리: infeasible 방지
        try:
            _sanitize_component_bounds(network)
        except Exception as _e_s:
            print(f"경계값 정리 경고: {_e_s}")
        return network
        
    except Exception as e:
        print(f"네트워크 생성 중 오류 발생: {str(e)}")
        traceback.print_exc()
        return None

def optimize_network(network):
    """네트워크 최적화"""
    if network is None:
        print("네트워크가 생성되지 않았습니다.")
        return False
    
    try:
        print("\n최적화 시작...")
        
        # CPU 코어 수 확인
        import multiprocessing
        num_cores = multiprocessing.cpu_count()
        print(f"사용 가능한 CPU 코어 수: {num_cores}")
        
        # 시점별 CHP 링크 요약(진단용)
        try:
            if not network.links.empty:
                chp_like = []
                for lk in network.links.index:
                    try:
                        b0 = str(network.links.at[lk, 'bus0'])
                        b1 = str(network.links.at[lk, 'bus1'])
                        b2 = str(network.links.at[lk, 'bus2']) if 'bus2' in network.links.columns and not pd.isna(network.links.at[lk, 'bus2']) else None
                        c0 = str(network.buses.at[b0, 'carrier']).lower() if b0 in network.buses.index else ''
                        c1 = str(network.buses.at[b1, 'carrier']).lower() if b1 in network.buses.index else ''
                        c2 = str(network.buses.at[b2, 'carrier']).lower() if (b2 and b2 in network.buses.index) else ''
                        if ('gas' in c0) and ('electric' in c1) and ('heat' in c2):
                            chp_like.append((lk, b0, b1, b2, float(network.links.at[lk, 'p_nom'])))
                    except Exception:
                        continue
                if chp_like:
                    print(f"CHP-유사 링크 {len(chp_like)}개 발견:")
                    for row in chp_like[:10]:
                        print(f"  {row[0]}: {row[1]} -> {row[2]}(EL), bus2={row[3]}(H), p_nom={row[4]}")
        except Exception:
            pass
        
        # 최적화 옵션 세트(순차 폴백)
        option_variants = [
            {'name': 'barrier',      'opts': {'threads': num_cores, 'lpmethod': 4, 'parallel': 1, 'barrier.algorithm': 3}},
            {'name': 'dual-simplex', 'opts': {'threads': num_cores, 'lpmethod': 2, 'parallel': 1}},
            {'name': 'primal-simplex','opts': {'threads': num_cores, 'lpmethod': 1, 'parallel': 1}}
        ]
        
        last_status = None
        last_error = None
        for variant in option_variants:
            vname = variant['name']
            sopts = variant['opts']
            print(f"\n[시도] CPLEX 방법: {vname}, 옵션: {sopts}")
            try:
                status = network.optimize(solver_name='cplex', solver_options=sopts)
                print(f"→ 상태: {status}")
                last_status = status
                if isinstance(status, tuple):
                    st_main = status[0]
                else:
                    st_main = str(status)
                if st_main and ('ok' in st_main.lower() or 'optimal' in st_main.lower()):
                    break
            except ValueError as e:
                if 'No objects to concatenate' in str(e):
                    print("경고: AC 각도 결과(v_ang)가 없어 후처리에서 concat 실패. 각도 결과 없이 계속 진행합니다.")
                    network.buses_t.v_ang = pd.DataFrame(index=network.snapshots, columns=network.buses.index)
                    last_status = 'ok'
                    break
                else:
                    last_error = str(e)
                    print(f"→ 예외: {last_error}")
                    continue
            except Exception as e:
                last_error = str(e)
                print(f"→ 예외: {last_error}")
                continue
        
        print(f"\n최종 최적화 상태: {last_status}")
        if hasattr(network, 'objective'):
            print(f"목적함수 값: {network.objective}")
        
        # 실패 시 LP 문제 내보내기(환경변수로 활성화)
        try:
            if (not last_status) or (isinstance(last_status, tuple) and all(x and ('unknown' in str(x).lower() or 'infeasible' in str(x).lower()) for x in last_status)) or ('unknown' in str(last_status).lower()):
                if os.environ.get('EXPORT_LP', '0') == '1':
                    export_dir = os.path.join('results', 'debug')
                    os.makedirs(export_dir, exist_ok=True)
                    if hasattr(network, 'model') and hasattr(network.model, 'to_file'):
                        lp_path = os.path.join(export_dir, 'failed_model.lp')
                        try:
                            network.model.to_file(lp_path)
                            print(f"실패 모델 LP 내보냄: {lp_path}")
                        except Exception as _e_lp:
                            print(f"LP 내보내기 실패: {_e_lp}")
                    else:
                        print("네트워크 모델 객체가 없어 LP 내보내기 불가")
        except Exception:
            pass
        
        return bool(last_status) and ('unknown' not in str(last_status).lower())
        
    except Exception as e:
        print(f"\n최적화 중 오류 발생: {str(e)}")
        traceback.print_exc()
        return False

def extract_results(network):
    """주요 결과 추출"""
    
    results = {
        'generator_output': network.generators_t.p,
        'node_prices': network.buses_t.marginal_price,
        'line_flows': network.lines_t.p0,
        'total_cost': network.objective,
        'load_balance': network.buses_t.p,
        'storage_state': network.storage_units_t.state_of_charge if not network.storage_units_t.empty else None
    }
    
    return results

def _classify_technology(gen_or_link_name):
    name = str(gen_or_link_name).strip().lower()
    if 'nuclear' in name:
        return '원자력'
    if 'pv' in name or 'solar' in name:
        return '태양광'
    if 'wt' in name or 'wind' in name:
        return '풍력'
    if 'hydro' in name or 'water' in name:
        return '수력'
    if 'coal' in name:
        return '석탄'
    if 'lng' in name or 'gas' in name:
        return 'LNG'
    if 'oil' in name or 'diesel' in name:
        return '석유'
    if 'biomass' in name or 'bio' in name:
        return '바이오'
    if 'geothermal' in name:
        return '지열'
    if 'chp' in name:
        return 'CHP'
    if 'heatpump' in name or 'heat_pump' in name or name.startswith('hp') or ' hp ' in name:
        return '히트펌프'
    if 'electrolyser' in name or 'electrolyzer' in name or 'electrolysis' in name:
        return '전해조'
    if 'h2' in name or 'hydrogen' in name:
        return '수소'
    if '_h_fallback_gen' in name or 'heat' in name:
        return '열'
    return '기타'

def _map_final_energy_from_carrier(carrier_value):
    c = str(carrier_value).strip().lower()
    if c in ['el', '전력', 'ac', 'dc'] or ('electric' in c or 'power' in c or 'hvac' in c or 'hvdc' in c):
        return '전력'
    if c in ['h', '열'] or ('heat' in c):
        return '열'
    if c in ['h2', '수소'] or ('hydrogen' in c):
        return '수소'
    return None

def build_final_energy_supply_tables(network):
    import pandas as _pd
    import numpy as _np
    rows_total = []
    rows_region = []
    bus_to_carrier = network.buses.carrier.to_dict() if not network.buses.empty else {}
    if not network.generators.empty and not network.generators_t.p.empty:
        for gen in network.generators.index:
            bus = network.generators.at[gen, 'bus']
            final_energy = _map_final_energy_from_carrier(bus_to_carrier.get(bus, ''))
            if final_energy is None:
                continue
            tech = _classify_technology(gen)
            series = network.generators_t.p[gen] if gen in network.generators_t.p.columns else _pd.Series(0.0, index=network.snapshots)
            val = float(_np.nansum(series.values))
            if val <= 0:
                continue
            rows_total.append([final_energy, tech, val])
            region = gen.split('_')[0] if '_' in gen else ''
            rows_region.append([region, final_energy, tech, val])
    def _add_link_port_supply(port_idx):
        p_attr = f"p{port_idx}"
        if not hasattr(network.links_t, p_attr):
            return None
        df_p = getattr(network.links_t, p_attr)
        if df_p is None or df_p.empty:
            return None
        bus_col = f"bus{port_idx}"
        for link in network.links.index:
            if bus_col not in network.links.columns:
                continue
            if _pd.isna(network.links.at[link, bus_col]):
                continue
            dest_bus = network.links.at[link, bus_col]
            final_energy = _map_final_energy_from_carrier(bus_to_carrier.get(dest_bus, ''))
            if final_energy is None:
                continue
            if link not in df_p.columns:
                continue
            series = df_p[link]
            delivered = (-series).clip(lower=0.0)
            val = float(_np.nansum(delivered.values))
            if val <= 0:
                continue
            tech = _classify_technology(link)
            rows_total.append([final_energy, tech, val])
            region = str(dest_bus).split('_')[0] if '_' in str(dest_bus) else ''
            rows_region.append([region, final_energy, tech, val])
    for k in [1, 2, 3]:
        _add_link_port_supply(k)
    total_df = _pd.DataFrame(rows_total, columns=['final_energy', 'technology', 'supply_MWh'])
    if not total_df.empty:
        total_df = total_df.groupby(['final_energy', 'technology'], as_index=False)['supply_MWh'].sum()
    by_region_df = _pd.DataFrame(rows_region, columns=['region', 'final_energy', 'technology', 'supply_MWh'])
    if not by_region_df.empty:
        by_region_df = by_region_df.groupby(['region', 'final_energy', 'technology'], as_index=False)['supply_MWh'].sum()
    return total_df, by_region_df

# 국가 기준 시간별 수급표(전력/열/수소) 생성

def _build_country_timeseries_tables(network):
    idx = network.snapshots
    zeros = pd.Series(0.0, index=idx)
    bus_to_carrier = network.buses.carrier.to_dict() if not network.buses.empty else {}

    # loads_t.p가 없거나 비어있으면 loads_t.p_set을 폴백으로 사용
    loads_p_df = None
    try:
        if hasattr(network.loads_t, 'p') and (network.loads_t.p is not None) and (not network.loads_t.p.empty):
            loads_p_df = network.loads_t.p
        elif hasattr(network.loads_t, 'p_set') and (network.loads_t.p_set is not None) and (not network.loads_t.p_set.empty):
            loads_p_df = network.loads_t.p_set
    except Exception:
        try:
            loads_p_df = getattr(network.loads_t, 'p_set')
        except Exception:
            loads_p_df = None

    def _fe(carrier):
        c = str(carrier).strip().lower()
        if c in ['el', '전력', 'ac', 'dc'] or ('electric' in c or 'power' in c or 'hvac' in c or 'hvdc' in c):
            return 'EL'
        if c in ['h', '열'] or ('heat' in c):
            return 'H'
        if c in ['h2', '수소'] or ('hydrogen' in c):
            return 'H2'
        return None

    def _fe_of_bus(bus_name):
        # 1) carrier 우선
        c = bus_to_carrier.get(bus_name, '')
        t = _fe(c)
        if t is not None:
            return t
        # 2) 버스명 토큰으로 판별 (예: 'SEL_EL', 'BSN_H2')
        try:
            s = str(bus_name)
            tokens = [tok for tok in s.split('_') if tok]
            if tokens:
                last = tokens[-1].strip().upper()
                if last in ['EL', 'H', 'H2']:
                    return last
        except Exception:
            pass
        return None

    def _accumulate_supply_and_load(target_type):
        supply = zeros.copy()
        load = zeros.copy()
        # 발전기 → 대상 최종단 버스에 연결된 출력 합계
        if (not network.generators.empty) and hasattr(network.generators_t, 'p') and (not network.generators_t.p.empty):
            for gen in network.generators.index:
                try:
                    bus = network.generators.at[gen, 'bus']
                    if _fe_of_bus(bus) != target_type:
                        continue
                    if gen in network.generators_t.p.columns:
                        supply = supply.add(network.generators_t.p[gen].fillna(0.0), fill_value=0.0)
                except Exception:
                    continue
        # 부하 → 대상 최종단 버스 부하 합계 (p가 없으면 p_set 사용)
        if (not network.loads.empty) and (loads_p_df is not None) and (not loads_p_df.empty):
            for ld in network.loads.index:
                try:
                    bus = network.loads.at[ld, 'bus']
                    if _fe_of_bus(bus) != target_type:
                        continue
                    if ld in loads_p_df.columns:
                        # p열이 전부 NaN이면 p_set으로 폴백
                        col_series = loads_p_df[ld]
                        if col_series.isna().all() and hasattr(network.loads_t, 'p_set') and (not network.loads_t.p_set.empty) and (ld in network.loads_t.p_set.columns):
                            col_series = network.loads_t.p_set[ld]
                        load = load.add(col_series.fillna(0.0), fill_value=0.0)
                except Exception:
                    continue
        # 링크 출력 포트에서 대상 최종단으로의 양(+) 유입을 공급으로 가정
        for k in [1, 2, 3]:
            p_attr = f"p{k}"
            if not hasattr(network.links_t, p_attr):
                continue
            dfp = getattr(network.links_t, p_attr)
            if dfp is None or dfp.empty:
                continue
            bus_col = f"bus{k}"
            if bus_col not in network.links.columns:
                continue
            for link in network.links.index:
                try:
                    if pd.isna(network.links.at[link, bus_col]):
                        continue
                    dest_bus = network.links.at[link, bus_col]
                    if _fe_of_bus(dest_bus) != target_type:
                        continue
                    if link not in dfp.columns:
                        continue
                    # 링크 출력 포트는 모델 부호상 목적 버스로 음(-) 유입될 수 있어, 공급으로는 부호 반전 후 양수만 합산
                    delivered = (-dfp[link]).fillna(0.0).clip(lower=0.0)
                    supply = supply.add(delivered, fill_value=0.0)
                except Exception:
                    continue
        return supply, load

    el_sup, el_load = _accumulate_supply_and_load('EL')
    h_sup, h_load = _accumulate_supply_and_load('H')
    h2_sup, h2_load = _accumulate_supply_and_load('H2')

    ts_el = pd.DataFrame({'Supply_MW': el_sup, 'Load_MW': el_load})
    ts_el['Net_MW'] = ts_el['Supply_MW'] - ts_el['Load_MW']

    ts_h = pd.DataFrame({'Supply_MW': h_sup, 'Load_MW': h_load})
    ts_h['Net_MW'] = ts_h['Supply_MW'] - ts_h['Load_MW']

    ts_h2 = pd.DataFrame({'Supply_MW': h2_sup, 'Load_MW': h2_load})
    ts_h2['Net_MW'] = ts_h2['Supply_MW'] - ts_h2['Load_MW']

    ts_el = ts_el.reindex(idx).fillna(0.0)
    ts_h = ts_h.reindex(idx).fillna(0.0)
    ts_h2 = ts_h2.reindex(idx).fillna(0.0)

    return ts_el, ts_h, ts_h2

def save_results(network, filename=None, subdir=None):
    """최적화 결과를 Excel 파일로 저장"""
    try:
        has_objective = hasattr(network, 'objective') and (network.objective is not None)
        if not has_objective:
            print("경고: 최적화 목적함수가 없지만, 중간 결과와 시계열을 저장합니다.")

        # 현재 시간을 포함한 폴더명 생성
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        if subdir:
            results_dir = subdir
            os.makedirs(results_dir, exist_ok=True)
        else:
            results_dir = f'results/{current_time}'
            os.makedirs('results', exist_ok=True)
            os.makedirs(results_dir, exist_ok=True)

        print(f"결과를 '{results_dir}' 폴더에 저장 중...")

        # 1. 기본 Excel 결과 파일
        excel_filename = f'{results_dir}/optimization_result_{current_time}.xlsx'
        with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
            # 발전기 출력 결과 (없으면 빈 프레임 저장)
            try:
                gtp = network.generators_t.p
            except Exception:
                gtp = pd.DataFrame()
            gtp.to_excel(writer, sheet_name='Generator_Output')

            # AC 선로 조류 결과
            try:
                if hasattr(network.lines_t, 'p0') and (not network.lines_t.p0.empty):
                    network.lines_t.p0.to_excel(writer, sheet_name='Line_Flow')
            except Exception:
                pass

            # HVDC Link 조류 결과 추가
            try:
                if hasattr(network.links_t, 'p0') and (not network.links_t.p0.empty):
                    network.links_t.p0.to_excel(writer, sheet_name='Link_Flow')
            except Exception:
                pass

            # 버스 정보
            try:
                bus_results = pd.DataFrame({
                    'v_nom': network.buses.v_nom,
                    'carrier': network.buses.carrier
                })
                bus_results.to_excel(writer, sheet_name='Bus_Info')
            except Exception:
                pass

            # 발전기 정보
            try:
                gen_results = pd.DataFrame({
                    'bus': network.generators.bus,
                    'p_nom': network.generators.p_nom,
                    'p_nom_min': network.generators.p_nom_min if 'p_nom_min' in network.generators.columns else pd.Series(index=network.generators.index, dtype=float),
                    'p_nom_extendable': network.generators.p_nom_extendable if 'p_nom_extendable' in network.generators.columns else pd.Series(index=network.generators.index, dtype=bool),
                    'p_max_pu': network.generators.p_max_pu,
                    'marginal_cost': network.generators.marginal_cost
                })
                if 'p_nom_opt' in network.generators.columns:
                    gen_results['p_nom_opt'] = network.generators.p_nom_opt
                gen_results.to_excel(writer, sheet_name='Generator_Info')
            except Exception:
                pass

            # Link 정보 추가
            try:
                if not network.links.empty:
                    link_results = pd.DataFrame({
                        'bus0': network.links.bus0,
                        'bus1': network.links.bus1,
                        'p_nom': network.links.p_nom,
                        'efficiency': network.links.efficiency
                    })
                    if 'p_nom_opt' in network.links.columns:
                        link_results['p_nom_opt'] = network.links.p_nom_opt
                    link_results.to_excel(writer, sheet_name='Link_Info')
            except Exception:
                pass

            # ESS 충방전 결과 (있는 경우에만)
            try:
                if hasattr(network, 'stores_t') and (not network.stores_t.p.empty):
                    network.stores_t.p.to_excel(writer, sheet_name='Storage_Power')
                if hasattr(network, 'stores_t') and (hasattr(network.stores_t, 'e') and (not network.stores_t.e.empty)):
                    network.stores_t.e.to_excel(writer, sheet_name='Storage_Energy')
            except Exception:
                pass

            # ESS 정보 (있는 경우에만)
            try:
                if not network.stores.empty:
                    store_results = pd.DataFrame({
                        'bus': network.stores.bus,
                        'carrier': network.stores.carrier,
                        'e_nom': network.stores.e_nom,
                        'e_cyclic': network.stores.e_cyclic
                    })
                    if 'e_nom_opt' in network.stores.columns:
                        store_results['e_nom_opt'] = network.stores.e_nom_opt
                    store_results.to_excel(writer, sheet_name='Storage_Info')
            except Exception:
                pass

            # 시간별 부하 결과 (p 없으면 p_set 저장)
            try:
                ltp = None
                if hasattr(network.loads_t, 'p') and (network.loads_t.p is not None) and (not network.loads_t.p.empty):
                    ltp = network.loads_t.p
                elif hasattr(network.loads_t, 'p_set') and (network.loads_t.p_set is not None) and (not network.loads_t.p_set.empty):
                    ltp = network.loads_t.p_set
                else:
                    ltp = pd.DataFrame(index=network.snapshots)
                ltp.to_excel(writer, sheet_name='Hourly_Loads')
            except Exception:
                pass

            # 최적화 요약
            try:
                status_label = 'Optimal' if has_objective else 'Infeasible/NoObjective'
                total_cost_val = float(network.objective) if has_objective else float('nan')
                summary = pd.DataFrame({
                    'Parameter': ['Total Cost', 'Status'],
                    'Value': [total_cost_val, status_label]
                })
                summary.to_excel(writer, sheet_name='Summary', index=False)
            except Exception:
                pass

            # 최종에너지별 공급 집계 시트 추가
            try:
                fe_total, fe_by_region = build_final_energy_supply_tables(network)
                if fe_total is not None and not fe_total.empty:
                    fe_total.to_excel(writer, sheet_name='FinalEnergy_Supply', index=False)
                if fe_by_region is not None and not fe_by_region.empty:
                    fe_by_region.to_excel(writer, sheet_name='FinalEnergy_Supply_ByRegion', index=False)
            except Exception as _e_fe:
                print(f"최종에너지 공급 집계 시트 저장 경고: {_e_fe}")

            # 국가 기준 시간별 수급표(전력/열/수소)
            try:
                ts_el, ts_h, ts_h2 = _build_country_timeseries_tables(network)
                ts_el.to_excel(writer, sheet_name='TS_Electricity_National')
                ts_h.to_excel(writer, sheet_name='TS_Heat_National')
                ts_h2.to_excel(writer, sheet_name='TS_Hydrogen_National')
            except Exception as _e_ts:
                print(f"국가 시간별 수급표 저장 경고: {_e_ts}")
                try:
                    # 폴백: 부하만이라도 기록
                    if hasattr(network.loads_t, 'p_set') and (network.loads_t.p_set is not None) and (not network.loads_t.p_set.empty):
                        network.loads_t.p_set.to_excel(writer, sheet_name='TS_Loads_Fallback')
                except Exception:
                    pass

        # 2. 개별 CSV 파일들 저장
        try:
            gen_info_df = pd.DataFrame({
                'bus': network.generators.bus,
                'p_nom': network.generators.p_nom,
                'p_nom_min': network.generators.p_nom_min if 'p_nom_min' in network.generators.columns else pd.Series(index=network.generators.index, dtype=float),
                'p_nom_extendable': network.generators.p_nom_extendable if 'p_nom_extendable' in network.generators.columns else pd.Series(index=network.generators.index, dtype=bool),
                'marginal_cost': network.generators.marginal_cost
            })
            if 'p_nom_opt' in network.generators.columns:
                gen_info_df['p_nom_opt'] = network.generators.p_nom_opt
            gen_info_df.to_csv(f'{results_dir}/optimization_result_{current_time}_generator_info.csv')
        except Exception as _e:
            print(f"Generator Info CSV 저장 경고: {_e}")

        # 발전기 출력
        try:
            network.generators_t.p.to_csv(f'{results_dir}/optimization_result_{current_time}_generator_output.csv')
        except Exception:
            pass

        # 부하
        try:
            if hasattr(network.loads_t, 'p') and (not network.loads_t.p.empty):
                network.loads_t.p.to_csv(f'{results_dir}/optimization_result_{current_time}_load.csv')
            elif hasattr(network.loads_t, 'p_set') and (not network.loads_t.p_set.empty):
                network.loads_t.p_set.to_csv(f'{results_dir}/optimization_result_{current_time}_load.csv')
        except Exception:
            pass

        # 저장장치 (있는 경우)
        try:
            if hasattr(network, 'stores_t') and not network.stores_t.p.empty:
                network.stores_t.p.to_csv(f'{results_dir}/optimization_result_{current_time}_storage.csv')
        except Exception:
            pass

        # 선로 사용량 (있는 경우)
        try:
            if not network.lines_t.p0.empty:
                network.lines_t.p0.to_csv(f'{results_dir}/optimization_result_{current_time}_line_usage.csv')
        except Exception:
            pass

        # 최종에너지 집계 CSV
        try:
            fe_total, fe_by_region = build_final_energy_supply_tables(network)
            if fe_total is not None and not fe_total.empty:
                fe_total.to_csv(f'{results_dir}/optimization_result_{current_time}_final_energy_supply.csv', index=False)
            if fe_by_region is not None and not fe_by_region.empty:
                fe_by_region.to_csv(f'{results_dir}/optimization_result_{current_time}_final_energy_supply_by_region.csv', index=False)
        except Exception as _e2:
            print(f"최종에너지 공급 집계 CSV 저장 경고: {_e2}")

        # 3. 통계 정보 JSON 파일
        try:
            total_cost_val = float(network.objective)
        except Exception:
            total_cost_val = float('nan')
        stats = {
            'timestamp': current_time,
            'total_cost': total_cost_val,
            'total_generators': len(network.generators),
            'total_buses': len(network.buses),
            'total_loads': len(network.loads),
            'total_stores': len(network.stores),
            'total_links': len(network.links)
        }

        import json
        with open(f'{results_dir}/optimization_result_{current_time}_stats.json', 'w', encoding='utf-8') as f:
            json.dump(stats, f, ensure_ascii=False, indent=2)

        # 4. PyPSA 네트워크 파일 저장
        network.export_to_netcdf(f'{results_dir}/optimization_result_{current_time}.nc')

        # 5. 지역별 분석 결과 생성
        analyze_regional_results(network, results_dir, current_time)

        # 6. 시각화 결과 생성
        create_visualizations(network, results_dir, current_time)

        print(f"결과가 '{results_dir}' 폴더에 저장되었습니다.")
        print(f"- Excel 파일: {excel_filename}")
        print(f"- CSV 파일들: generator_output, load, storage, line_usage")
        print(f"- 통계 파일: stats.json")
        print(f"- 네트워크 파일: .nc")
        print(f"- 시각화 파일들: PNG, HTML")

        return True

    except Exception as e:
        print(f"결과 저장 중 오류 발생: {str(e)}")
        traceback.print_exc()
        return False

def create_visualizations(network, results_dir, current_time):
    if os.environ.get('DISABLE_PLOTS','0') == '1':
        print('시각화 생략(DISABLE_PLOTS=1)')
        return True
    """시각화 결과 생성"""
    try:
        import matplotlib.pyplot as plt
        import plotly.graph_objects as go
        import plotly.express as px
        from plotly.subplots import make_subplots
        import plotly.offline as pyo
        
        # 한글 폰트 설정(윈도우: 맑은 고딕)
        try:
            import matplotlib as mpl
            mpl.rcParams['font.family'] = ['Malgun Gothic', 'DejaVu Sans']
            mpl.rcParams['axes.unicode_minus'] = False
            import warnings as _warn
            import logging as _logging
            _warn.filterwarnings('ignore', message='findfont:', category=UserWarning, module='matplotlib')
            try:
                _logging.getLogger('matplotlib.font_manager').setLevel(_logging.ERROR)
            except Exception:
                pass
        except Exception:
            pass
        
        print("시각화 결과 생성 중...")
        
        # 1. 지역별 에너지 밸런스 차트
        gen_output = network.generators_t.p
        load_output = network.loads_t.p
        
        # 지역별 발전량과 부하량 계산
        regional_generation = {}
        regional_load = {}
        
        for gen_name in gen_output.columns:
            if '_' in gen_name:
                region = gen_name.split('_')[0]
                if region not in regional_generation:
                    regional_generation[region] = 0
                regional_generation[region] += gen_output[gen_name].sum()
        
        for load_name in load_output.columns:
            if '_' in load_name:
                region = load_name.split('_')[0]
                if region not in regional_load:
                    regional_load[region] = 0
                regional_load[region] += load_output[load_name].sum()
        
        # 지역별 에너지 밸런스 DataFrame 생성
        regions = list(set(list(regional_generation.keys()) + list(regional_load.keys())))
        balance_data = []
        for region in regions:
            gen = regional_generation.get(region, 0)
            load = regional_load.get(region, 0)
            balance = gen - load
            balance_data.append([region, gen, load, balance])
        
        balance_df = pd.DataFrame(balance_data, columns=['지역', '발전량(MWh)', '부하량(MWh)', '밸런스(MWh)'])
        balance_df.to_csv(f'{results_dir}/regional_energy_balance.csv', index=False, encoding='utf-8-sig')
        
        # 지역별 에너지 밸런스 차트 생성
        fig, ax = plt.subplots(figsize=(15, 8))
        x = range(len(balance_df))
        width = 0.35
        
        ax.bar([i - width/2 for i in x], balance_df['발전량(MWh)'], width, label='발전량', alpha=0.8)
        ax.bar([i + width/2 for i in x], balance_df['부하량(MWh)'], width, label='부하량', alpha=0.8)
        
        ax.set_xlabel('지역')
        ax.set_ylabel('에너지 (MWh)')
        ax.set_title('지역별 에너지 밸런스')
        ax.set_xticks(x)
        ax.set_xticklabels(balance_df['지역'], rotation=45)
        ax.legend()
        ax.grid(True, alpha=0.3)
        
        plt.tight_layout()
        plt.savefig(f'{results_dir}/regional_energy_balance.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        # 2. 지역별 재생에너지 비율 차트
        renewable_ratio_data = []
        for region in regions:
            total_gen = 0
            renewable_gen = 0
            
            for gen_name in gen_output.columns:
                if gen_name.startswith(region + '_'):
                    gen_sum = gen_output[gen_name].sum()
                    total_gen += gen_sum
                    if 'PV' in gen_name or 'WT' in gen_name or 'Wind' in gen_name or 'Solar' in gen_name:
                        renewable_gen += gen_sum
            
            if total_gen > 0:
                ratio = (renewable_gen / total_gen) * 100
            else:
                ratio = 0
            
            renewable_ratio_data.append([region, renewable_gen, total_gen, ratio])
        
        renewable_df = pd.DataFrame(renewable_ratio_data, columns=['지역', '재생에너지(MWh)', '총발전량(MWh)', '재생에너지비율(%)'])
        
        # 재생에너지 비율 차트
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(20, 8))
        
        # 재생에너지 비율 막대 차트
        ax1.bar(renewable_df['지역'], renewable_df['재생에너지비율(%)'], color='green', alpha=0.7)
        ax1.set_xlabel('지역')
        ax1.set_ylabel('재생에너지 비율 (%)')
        ax1.set_title('지역별 재생에너지 비율')
        ax1.tick_params(axis='x', rotation=45)
        ax1.grid(True, alpha=0.3)
        
        # 재생에너지 vs 총발전량 비교
        x = range(len(renewable_df))
        width = 0.35
        
        ax2.bar([i - width/2 for i in x], renewable_df['재생에너지(MWh)'], width, label='재생에너지', color='green', alpha=0.7)
        ax2.bar([i + width/2 for i in x], renewable_df['총발전량(MWh)'] - renewable_df['재생에너지(MWh)'], width, label='기타 발전', color='gray', alpha=0.7)
        
        ax2.set_xlabel('지역')
        ax2.set_ylabel('발전량 (MWh)')
        ax2.set_title('지역별 발전원별 발전량')
        ax2.set_xticks(x)
        ax2.set_xticklabels(renewable_df['지역'], rotation=45)
        ax2.legend()
        ax2.grid(True, alpha=0.3)
        
        plt.tight_layout()
        plt.savefig(f'{results_dir}/regional_renewable_ratio.png', dpi=300, bbox_inches='tight')
        plt.close()
        
        # 3. 송전선로 조류 지도 (HTML)
        if not network.lines_t.p0.empty:
            # 송전선로 조류 데이터 준비
            line_flows = network.lines_t.p0.mean()  # 평균 조류
            transmission_data = []
            
            for line_name in line_flows.index:
                if line_name in network.lines.index:
                    bus0 = network.lines.at[line_name, 'bus0']
                    bus1 = network.lines.at[line_name, 'bus1']
                    flow = line_flows[line_name]
                    capacity = network.lines.at[line_name, 's_nom']
                    utilization = abs(flow) / capacity * 100 if capacity > 0 else 0
                    
                    transmission_data.append([line_name, bus0, bus1, flow, capacity, utilization])
            
            transmission_df = pd.DataFrame(transmission_data, 
                                         columns=['선로명', '시작버스', '종료버스', '평균조류(MW)', '용량(MVA)', '이용률(%)'])
            transmission_df.to_csv(f'{results_dir}/transmission_flow.csv', index=False, encoding='utf-8-sig')
            
            # 송전선로 네트워크 그래프
            import networkx as nx
            
            G = nx.Graph()
            
            # 노드 추가 (버스)
            for bus in network.buses.index:
                G.add_node(bus)
            
            # 엣지 추가 (선로)
            for line_name in network.lines.index:
                bus0 = network.lines.at[line_name, 'bus0']
                bus1 = network.lines.at[line_name, 'bus1']
                if line_name in line_flows.index:
                    flow = abs(line_flows[line_name])
                    G.add_edge(bus0, bus1, weight=flow, line_name=line_name)
            
            # 네트워크 그래프 그리기
            plt.figure(figsize=(20, 15))
            pos = nx.spring_layout(G, k=3, iterations=50)
            
            # 노드 그리기
            nx.draw_networkx_nodes(G, pos, node_color='lightblue', node_size=500, alpha=0.8)
            
            # 엣지 그리기 (조류 크기에 따라 두께 조정)
            edges = G.edges()
            weights = [G[u][v]['weight'] for u, v in edges]
            max_weight = max(weights) if weights else 1e-6
            edge_widths = [w/max_weight * 5 + 0.5 for w in weights]
            
            nx.draw_networkx_edges(G, pos, width=edge_widths, alpha=0.6, edge_color='red')
            
            # 라벨 그리기
            nx.draw_networkx_labels(G, pos, font_size=8, font_weight='bold')
            
            plt.title('송전선로 네트워크 및 조류 현황', fontsize=16, fontweight='bold')
            plt.axis('off')
            plt.tight_layout()
            plt.savefig(f'{results_dir}/transmission_network_graph.png', dpi=300, bbox_inches='tight')
            plt.close()
            
            # 인터랙티브 송전선로 지도 (HTML)
            fig = go.Figure()
            
            # 선로별 조류 막대 차트
            fig.add_trace(go.Bar(
                x=transmission_df['선로명'],
                y=transmission_df['이용률(%)'],
                name='선로 이용률',
                text=transmission_df['평균조류(MW)'].round(1),
                textposition='auto',
                hovertemplate='<b>%{x}</b><br>이용률: %{y:.1f}%<br>조류: %{text} MW<extra></extra>'
            ))
            
            fig.update_layout(
                title='송전선로별 이용률 및 조류 현황',
                xaxis_title='송전선로',
                yaxis_title='이용률 (%)',
                hovermode='x unified',
                height=600
            )
            
            pyo.plot(fig, filename=f'{results_dir}/transmission_flow_map.html', auto_open=False)
        
        # 4. 한국 지도 기반 송전선로 시각화 추가
        try:
            create_korea_transmission_map(network, results_dir, current_time)
        except Exception as e:
            print(f"한국 지도 기반 송전선로 시각화 생성 중 오류: {str(e)}")
        
        # 5. 이전 버전의 한국 지도 시각화 추가
        try:
            create_legacy_korea_map(network, results_dir, current_time)
        except Exception as e:
            print(f"이전 버전 한국 지도 시각화 생성 중 오류: {str(e)}")
        
        print("시각화 결과 생성 완료:")
        print("- 지역별 에너지 밸런스 차트")
        print("- 지역별 재생에너지 비율 차트")
        print("- 송전선로 네트워크 그래프")
        print("- 인터랙티브 송전선로 지도")
        print("- 한국 지도 기반 송전선로 시각화")
        print("- 이전 버전 한국 지도 시각화")
        
        return True
        
    except Exception as e:
        print(f"시각화 생성 중 오류 발생: {str(e)}")
        traceback.print_exc()
        return False

def create_korea_transmission_map(network, results_dir, current_time):
    """한국 지도 기반 송전선로 시각화"""
    try:
        from korea_map import KoreaMapVisualizer
        import matplotlib.pyplot as plt
        
        print("한국 지도 기반 송전선로 시각화 생성 중...")
        
        # 지역 코드와 한국어 이름 매핑
        region_mapping = {
            'SEL': '서울특별시',
            'BSN': '부산광역시', 
            'DGU': '대구광역시',
            'ICN': '인천광역시',
            'GWJ': '광주광역시',
            'DJN': '대전광역시',
            'USN': '울산광역시',
            'SJG': '세종특별자치시',
            'GGD': '경기도',
            'GWD': '강원도',
            'CBD': '충청북도',
            'CND': '충청남도',
            'JBD': '전라북도',
            'JND': '전라남도',
            'GBD': '경상북도',
            'GND': '경상남도',
            'JJD': '제주특별자치도'
        }
        
        # 송전선로 조류 데이터 준비
        line_flows = network.lines_t.p0.mean() if not network.lines_t.p0.empty else pd.Series()
        
        # 송전선로 정보를 DataFrame으로 구성
        transmission_lines = []
        
        for line_name in network.lines.index:
            bus0 = network.lines.at[line_name, 'bus0']
            bus1 = network.lines.at[line_name, 'bus1']
            
            # 버스 이름에서 지역 코드 추출
            region0 = bus0.split('_')[0] if '_' in bus0 else bus0
            region1 = bus1.split('_')[0] if '_' in bus1 else bus1
            
            # 지역 코드를 한국어 이름으로 변환
            region0_name = region_mapping.get(region0, region0)
            region1_name = region_mapping.get(region1, region1)
            
            # 조류 정보
            flow = line_flows.get(line_name, 0) if line_name in line_flows.index else 0
            capacity = network.lines.at[line_name, 's_nom']
            utilization = abs(flow) / capacity * 100 if capacity > 0 else 0
            
            transmission_lines.append({
                '선로명': line_name,
                '시작지역': region0_name,
                '종료지역': region1_name,
                '조류(MW)': flow,
                '용량(MVA)': capacity,
                '이용률(%)': utilization
            })
        
        # DataFrame 생성
        transmission_df = pd.DataFrame(transmission_lines)
        
        # 한국 지도 시각화 객체 생성
        visualizer = KoreaMapVisualizer()
        
        if visualizer.load_map_data():
            # 기본 한국 지도 생성
            map_save_path = f'{results_dir}/korea_transmission_map_{current_time}.png'
            
            # 지도 그리기
            fig, ax = plt.subplots(figsize=(15, 12))
            
            # 행정구역 경계 그리기
            visualizer.map_data.plot(ax=ax, 
                                   color='lightgray', 
                                   edgecolor='black', 
                                   linewidth=0.5,
                                   alpha=0.7)
            
            # 지역 중심점 계산
            region_centroids = {}
            for idx, row in visualizer.map_data.iterrows():
                region_name = row['SIDO_NM']
                centroid = row.geometry.centroid
                region_centroids[region_name] = (centroid.x, centroid.y)
            
            # 송전선로 그리기
            for _, line in transmission_df.iterrows():
                start_region = line['시작지역']
                end_region = line['종료지역']
                utilization = line['이용률(%)']
                
                if start_region in region_centroids and end_region in region_centroids:
                    start_x, start_y = region_centroids[start_region]
                    end_x, end_y = region_centroids[end_region]
                    
                    # 이용률에 따른 선 두께와 색상 설정
                    line_width = max(0.5, utilization / 20)  # 최소 0.5, 최대 5
                    
                    if utilization > 80:
                        color = 'red'
                    elif utilization > 60:
                        color = 'orange'
                    elif utilization > 40:
                        color = 'yellow'
                    else:
                        color = 'green'
                    
                    # 송전선로 그리기
                    ax.plot([start_x, end_x], [start_y, end_y], 
                           color=color, linewidth=line_width, alpha=0.8)
                    
                    # 중점에 이용률 표시 (이용률이 높은 경우만)
                    if utilization > 50:
                        mid_x = (start_x + end_x) / 2
                        mid_y = (start_y + end_y) / 2
                        ax.text(mid_x, mid_y, f'{utilization:.0f}%', 
                               fontsize=8, ha='center', va='center',
                               bbox=dict(facecolor='white', alpha=0.7, edgecolor='none'))
            
            # 지역 중심점에 지역명 표시
            for region_name, (x, y) in region_centroids.items():
                # 지역 코드로 변환
                region_code = None
                for code, name in region_mapping.items():
                    if name == region_name:
                        region_code = code
                        break
                
                if region_code:
                    ax.plot(x, y, 'ko', markersize=8, alpha=0.8)
                    ax.text(x, y + 20000, region_code, fontsize=10, ha='center', va='bottom',
                           fontweight='bold',
                           bbox=dict(facecolor='white', alpha=0.8, edgecolor='black'))
            
            # 범례 추가
            legend_elements = [
                plt.Line2D([0], [0], color='green', lw=2, label='이용률 < 40%'),
                plt.Line2D([0], [0], color='yellow', lw=2, label='이용률 40-60%'),
                plt.Line2D([0], [0], color='orange', lw=2, label='이용률 60-80%'),
                plt.Line2D([0], [0], color='red', lw=2, label='이용률 > 80%')
            ]
            ax.legend(handles=legend_elements, loc='upper left', bbox_to_anchor=(0.02, 0.98))
            
            ax.set_title('한국 송전선로 이용률 현황', fontsize=16, fontweight='bold', pad=20)
            ax.set_axis_off()
            
            plt.tight_layout()
            plt.savefig(map_save_path, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()
            
            print(f"한국 지도 기반 송전선로 시각화가 '{map_save_path}'에 저장되었습니다.")
            
            # 송전선로 상세 정보 CSV 저장
            transmission_df.to_csv(f'{results_dir}/korea_transmission_details_{current_time}.csv', 
                                 index=False, encoding='utf-8-sig')
            
            return True
        
            print("한국 지도 데이터를 로드할 수 없습니다.")
            return False
            
    except Exception as e:
        print(f"한국 지도 기반 송전선로 시각화 생성 중 오류: {str(e)}")
        traceback.print_exc()
        return False

def create_legacy_korea_map(network, results_dir, current_time):
    """이전 버전의 한국 지도 시각화 (간단한 지역 연결 지도)"""
    try:
        import matplotlib.pyplot as plt
        import matplotlib.patches as patches
        from matplotlib.patches import FancyBboxPatch
        import numpy as np
        
        print("이전 버전 한국 지도 시각화 생성 중...")
        
        # 한국 지역 좌표 정의 (간단한 버전)
        region_coords = {
            'SEL': (5, 7),    # 서울
            'ICN': (4, 7),    # 인천
            'GGD': (5, 6),    # 경기
            'GWD': (7, 8),    # 강원
            'CBD': (6, 5),    # 충북
            'CND': (4, 5),    # 충남
            'DJN': (5, 4),    # 대전
            'SJG': (5, 4.5),  # 세종
            'JBD': (3, 3),    # 전북
            'JND': (2, 2),    # 전남
            'GWJ': (3, 2.5),  # 광주
            'GBD': (7, 4),    # 경북
            'DGU': (7, 3),    # 대구
            'GND': (6, 2),    # 경남
            'BSN': (7, 1),    # 부산
            'USN': (7.5, 1.5), # 울산
            'JJD': (1, 0)     # 제주
        }
        
        # 송전선로 조류 데이터 준비
        line_flows = network.lines_t.p0.mean() if not network.lines_t.p0.empty else pd.Series()
        
        # 지역별 발전량 계산
        gen_output = network.generators_t.p
        regional_generation = {}
        regional_renewable = {}
        
        for gen_name in gen_output.columns:
            if '_' in gen_name:
                region = gen_name.split('_')[0]
                gen_sum = gen_output[gen_name].sum()
                
                if region not in regional_generation:
                    regional_generation[region] = 0
                    regional_renewable[region] = 0
                
                regional_generation[region] += gen_sum
                
                # 재생에너지 여부 확인
                if any(keyword in gen_name for keyword in ['PV', 'WT', 'Solar', 'Wind']):
                    regional_renewable[region] += gen_sum
        
        # 지역별 부하량 계산
        load_output = network.loads_t.p
        regional_load = {}
        
        for load_name in load_output.columns:
            if '_' in load_name:
                region = load_name.split('_')[0]
                load_sum = load_output[load_name].sum()
                
                if region not in regional_load:
                    regional_load[region] = 0
                regional_load[region] += load_sum
        
        # 그래프 생성
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(20, 10))
        
        # 첫 번째 지도: 발전량과 송전선로
        ax1.set_xlim(0, 9)
        ax1.set_ylim(-1, 9)
        ax1.set_aspect('equal')
        
        # 지역별 발전량을 원의 크기로 표현
        max_generation = max(regional_generation.values()) if regional_generation else 1e-6
        
        for region, coords in region_coords.items():
            x, y = coords
            generation = regional_generation.get(region, 0)
            renewable = regional_renewable.get(region, 0)
            load = regional_load.get(region, 0)
            
            # 발전량에 비례한 원의 크기
            size = max(50, (generation / max_generation) * 1000) if max_generation > 0 else 50
            
            # 재생에너지 비율에 따른 색상
            renewable_ratio = renewable / generation if generation > 0 else 0
            color = plt.cm.RdYlGn(renewable_ratio)
            
            # 지역 원 그리기
            circle = plt.Circle((x, y), np.sqrt(size)/20, color=color, alpha=0.7, edgecolor='black')
            ax1.add_patch(circle)
            
            # 지역명 표시
            ax1.text(x, y, region, ha='center', va='center', fontsize=8, fontweight='bold')
            
            # 발전량 정보 표시
            info_text = f"{generation/1000:.0f}GWh"
            if renewable_ratio > 0:
                info_text += f"\n재생{renewable_ratio*100:.0f}%"
            ax1.text(x, y-0.5, info_text, ha='center', va='top', fontsize=6)
        
        # 송전선로 그리기
        for line_name in network.lines.index:
            bus0 = network.lines.at[line_name, 'bus0']
            bus1 = network.lines.at[line_name, 'bus1']
            
            # 버스 이름에서 지역 코드 추출
            region0 = bus0.split('_')[0] if '_' in bus0 else bus0
            region1 = bus1.split('_')[0] if '_' in bus1 else bus1
            
            if region0 in region_coords and region1 in region_coords:
                x0, y0 = region_coords[region0]
                x1, y1 = region_coords[region1]
                
                # 조류 정보
                flow = abs(line_flows.get(line_name, 0)) if line_name in line_flows.index else 0
                capacity = network.lines.at[line_name, 's_nom']
                utilization = flow / capacity * 100 if capacity > 0 else 0
                
                # 이용률에 따른 선 색상과 두께
                if utilization > 80:
                    color = 'red'
                    width = 3
                elif utilization > 60:
                    color = 'orange'
                    width = 2.5
                elif utilization > 40:
                    color = 'yellow'
                    width = 2
                else:
                    color = 'green'
                    width = 1.5
                
                # 송전선로 그리기
                ax1.plot([x0, x1], [y0, y1], color=color, linewidth=width, alpha=0.8)
                
                # 이용률이 높은 경우 수치 표시
                if utilization > 50:
                    mid_x, mid_y = (x0 + x1) / 2, (y0 + y1) / 2
                    ax1.text(mid_x, mid_y, f'{utilization:.0f}%', 
                            fontsize=6, ha='center', va='center',
                            bbox=dict(facecolor='white', alpha=0.7, edgecolor='none', pad=1))
        
        ax1.set_title('지역별 발전량 및 송전선로 이용률', fontsize=14, fontweight='bold')
        ax1.set_xlabel('재생에너지 비율: 빨간색(낮음) → 녹색(높음)', fontsize=10)
        ax1.grid(True, alpha=0.3)
        ax1.set_xticks([])
        ax1.set_yticks([])
        
        # 두 번째 지도: 에너지 밸런스
        ax2.set_xlim(0, 9)
        ax2.set_ylim(-1, 9)
        ax2.set_aspect('equal')
        
        # 지역별 에너지 밸런스 (발전량 - 부하량)
        for region, coords in region_coords.items():
            x, y = coords
            generation = regional_generation.get(region, 0)
            load = regional_load.get(region, 0)
            balance = generation - load
            
            # 밸런스에 따른 색상 (잉여: 파란색, 부족: 빨간색)
            gen_max = max([1e-6] + list(regional_generation.values()))
            load_max = max([1e-6] + list(regional_load.values()))
            if balance > 0:
                color = 'blue'
                alpha = min(0.8, abs(balance) / gen_max * 2)
            else:
                color = 'red'
                alpha = min(0.8, abs(balance) / load_max * 2)
            
            # 밸런스 크기에 비례한 원
            denom = max([1e-6] + list(regional_generation.values()) + list(regional_load.values()))
            size = max(50, abs(balance) / denom * 1000)
            
            circle = plt.Circle((x, y), np.sqrt(size)/20, color=color, alpha=alpha, edgecolor='black')
            ax2.add_patch(circle)
            
            # 지역명과 밸런스 정보
            ax2.text(x, y, region, ha='center', va='center', fontsize=8, fontweight='bold', color='white')
            balance_text = f"{balance/1000:.0f}GWh"
            if balance > 0:
                balance_text = "+" + balance_text
            ax2.text(x, y-0.5, balance_text, ha='center', va='top', fontsize=6)
        
        ax2.set_title('지역별 에너지 밸런스 (발전량 - 부하량)', fontsize=14, fontweight='bold')
        ax2.set_xlabel('파란색: 잉여, 빨간색: 부족', fontsize=10)
        ax2.grid(True, alpha=0.3)
        ax2.set_xticks([])
        ax2.set_yticks([])
        
        # 범례 추가
        legend_elements1 = [
            plt.Line2D([0], [0], color='green', lw=2, label='이용률 < 40%'),
            plt.Line2D([0], [0], color='yellow', lw=2, label='이용률 40-60%'),
            plt.Line2D([0], [0], color='orange', lw=2, label='이용률 60-80%'),
            plt.Line2D([0], [0], color='red', lw=2, label='이용률 > 80%')
        ]
        ax1.legend(handles=legend_elements1, loc='upper left')
        
        legend_elements2 = [
            plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='blue', markersize=10, label='에너지 잉여'),
            plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='red', markersize=10, label='에너지 부족')
        ]
        ax2.legend(handles=legend_elements2, loc='upper left')
        
        plt.tight_layout()
        
        # 파일 저장
        legacy_map_path = f'{results_dir}/legacy_korea_transmission_map_{current_time}.png'
        plt.savefig(legacy_map_path, dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()
        
        print(f"이전 버전 한국 지도 시각화가 '{legacy_map_path}'에 저장되었습니다.")
        
        # 상세 정보 CSV 저장
        legacy_data = []
        for region in region_coords.keys():
            generation = regional_generation.get(region, 0)
            renewable = regional_renewable.get(region, 0)
            load = regional_load.get(region, 0)
            balance = generation - load
            renewable_ratio = renewable / generation * 100 if generation > 0 else 0
            
            legacy_data.append([
                region, generation/1000, renewable/1000, load/1000, 
                balance/1000, renewable_ratio
            ])
        
        legacy_df = pd.DataFrame(legacy_data, columns=[
            '지역', '발전량(GWh)', '재생에너지(GWh)', '부하량(GWh)', 
            '밸런스(GWh)', '재생에너지비율(%)'
        ])
        legacy_df.to_csv(f'{results_dir}/legacy_korea_energy_summary_{current_time}.csv', 
                        index=False, encoding='utf-8-sig')
        
        return True
        
    except Exception as e:
        print(f"이전 버전 한국 지도 시각화 생성 중 오류: {str(e)}")
        traceback.print_exc()
        return False

def validate_input_data(input_data):
    """입력 데이터 유효성 검사"""
    required_sheets = ['buses', 'generators', 'lines', 'loads', 'timeseries']
    
    # 필수 시트 확인
    for sheet in required_sheets:
        if sheet not in input_data:
            raise ValueError(f"필수 시트 '{sheet}'가 없습니다.")
        if input_data[sheet].empty:
            raise ValueError(f"'{sheet}' 시트가 비어있습니다.")
    
    # 데이터 타입 확인 및 변환
    for _, row in input_data['buses'].iterrows():
        if not isinstance(row['name'], str):
            input_data['buses'].loc[_, 'name'] = str(row['name'])
    
    # timeseries 데이터 확인
    timeseries = input_data['timeseries'].iloc[0]
    if not pd.to_datetime(timeseries['start_time']):
        raise ValueError("잘못된 start_time 형식입니다.")
    if not pd.to_datetime(timeseries['end_time']):
        raise ValueError("잘못된 end_time 형식입니다.")
    if 'h' not in str(timeseries['frequency']).lower():
        raise ValueError("frequency는 'h' 형식이어야 합니다.")

    return input_data

def check_network_connections(network):
    """네트워크 연결 상태 확인"""
    print("\n=== 네트워크 연결 상태 확인 ===")
    
    # 1. 버스 연결 확인
    print("\n버스 정보:")
    for bus in network.buses.index:
        print(f"\n버스 {bus}:")
        # 연결된 발전기
        gens = network.generators[network.generators.bus == bus].index
        print(f"연결된 발전기: {list(gens)}")
        # 연결된 부하
        loads = network.loads[network.loads.bus == bus].index
        print(f"연결된 부하: {list(loads)}")
        # 연결된 저장장치
        stores = network.stores[network.stores.bus == bus].index
        print(f"연결된 저장장치: {list(stores)}")
    
    # 2. ESS 연결 및 설정 확인
    print("\nESS 상세 설정:")
    for store in network.stores.index:
        print(f"\n저장장치 {store}:")
        print(f"연결된 버스: {network.stores.at[store, 'bus']}")
        print(f"저장용량: {network.stores.at[store, 'e_nom']} MWh")
        print(f"충전효율: {network.stores.at[store, 'efficiency_store'] if 'efficiency_store' in network.stores else 1.0}")
        print(f"방전효율: {network.stores.at[store, 'efficiency_dispatch'] if 'efficiency_dispatch' in network.stores else 1.0}")
        print(f"순환 운전: {network.stores.at[store, 'e_cyclic']}")
        print(f"초기 충전상태: {network.stores.at[store, 'e_initial'] if 'e_initial' in network.stores else 0}")

    # 3. 선로 연결 확인
    print("\n선로 연결:")
    if not network.lines.empty:
        for line in network.lines.index:
            print(f"\n선로 {line}:")
            print(f"From: {network.lines.at[line, 'bus0']} To: {network.lines.at[line, 'bus1']}")
            print(f"용량: {network.lines.at[line, 's_nom']} MVA")

def check_excel_data_loading(input_data):
    """엑셀 데이터 로드 상태 확인"""
    print("\n=== 엑셀 데이터 로드 상태 확인 ===")
    
    # 1. 시트 존재 여부 확인
    print("\n로드된 시트:")
    for sheet_name in input_data.keys():
        print(f"\n[{sheet_name}] 시트:")
        print(f"행 수: {len(input_data[sheet_name])}")
        print(f"컬럼: {list(input_data[sheet_name].columns)}")
    
    # 2. stores 시트 상세 확인
    if 'stores' in input_data:
        print("\n\n=== Stores 시트 상세 데이터 ===")
        stores_df = input_data['stores']
        print("\n컬럼별 데이터 타입:")
        print(stores_df.dtypes)
        print("\n실제 데이터:")
        print(stores_df)
        
        # 필수 컬럼 확인
        required_columns = [
            'name', 'bus', 'carrier', 'e_nom', 'e_nom_extendable', 
            'e_cyclic', 'standing_loss', 'efficiency_store', 
            'efficiency_dispatch', 'e_initial', 'e_nom_max'
        ]
        missing_columns = [col for col in required_columns if col not in stores_df.columns]
        if missing_columns:
            print(f"\n누락된 필수 컬럼: {missing_columns}")
    
    # 3. links 시트 상세 확인
    if 'links' in input_data:
        print("\n\n=== Links 시트 상세 데이터 ===")
        links_df = input_data['links']
        print("\n컬럼별 데이터 타입:")
        print(links_df.dtypes)
        print("\n실제 데이터:")
        print(links_df)
        
        # 필수 컬럼 확인
        required_columns = [
            'name', 'bus0', 'bus1', 'efficiency', 'p_nom', 
            'p_nom_extendable', 'p_nom_max'
        ]
        missing_columns = [col for col in required_columns if col not in links_df.columns]
        if missing_columns:
            print(f"\n누락된 필수 컬럼: {missing_columns}")

def analyze_regional_results(network, results_dir, current_time):
    """지역별 분석 결과 생성"""
    try:
        print("지역별 분석 결과 생성 중...")
        
        # 발전기 출력 데이터
        gen_output = network.generators_t.p
        
        # 지역별 발전량 계산
        regional_generation = {}
        for gen_name in gen_output.columns:
            # 지역 코드 추출 (A_, B_ 등)
            if '_' in gen_name:
                region = gen_name.split('_')[0]
                if region not in regional_generation:
                    regional_generation[region] = 0
                regional_generation[region] += gen_output[gen_name].sum()
        
        # 지역별 발전량 DataFrame 생성
        regional_df = pd.DataFrame(list(regional_generation.items()), 
                                 columns=['지역', '총발전량(MWh)'])
        regional_df.to_csv(f'{results_dir}/optimization_result_{current_time}_지역별_발전량.csv', 
                          index=False, encoding='utf-8-sig')
        
        # 발전원별 발전량 계산
        generation_by_type = {}
        for gen_name in gen_output.columns:
            # 발전원 타입 추출(보강)
            gen_type = 'Unknown'
            lname = gen_name.lower()
            if 'nuclear' in lname:
                gen_type = '원자력'
            elif 'pv' in lname or 'solar' in lname:
                gen_type = '태양광'
            elif 'wt' in lname or 'wind' in lname:
                gen_type = '풍력'
            elif 'hydro' in lname or 'water' in lname:
                gen_type = '수력'
            elif 'coal' in lname:
                gen_type = '석탄'
            elif 'lng' in lname or 'gas' in lname:
                gen_type = 'LNG'
            elif 'oil' in lname or 'diesel' in lname:
                gen_type = '석유'
            elif 'biomass' in lname or 'bio' in lname:
                gen_type = '바이오'
            elif 'geothermal' in lname:
                gen_type = '지열'
            elif 'chp' in lname:
                gen_type = '열병합'
            elif 'h2' in lname or 'hydrogen' in lname:
                gen_type = '수소'
            elif '_h_fallback_gen' in lname or 'heat' in lname:
                gen_type = '열'
            
            if gen_type not in generation_by_type:
                generation_by_type[gen_type] = 0
            generation_by_type[gen_type] += gen_output[gen_name].sum()
        
        # 발전원별 발전량 DataFrame 생성
        type_df = pd.DataFrame(list(generation_by_type.items()), 
                             columns=['발전원', '총발전량(MWh)'])
        type_df.to_csv(f'{results_dir}/optimization_result_{current_time}_발전원별_발전량.csv', 
                      index=False, encoding='utf-8-sig')
        
        # 지역별 발전원별 발전량 계산
        regional_type_generation = {}
        for gen_name in gen_output.columns:
            if '_' in gen_name:
                region = gen_name.split('_')[0]
                
                # 발전원 타입 추출
                gen_type = 'Unknown'
                lname = gen_name.lower()
                if 'nuclear' in lname:
                    gen_type = '원자력'
                elif 'pv' in lname or 'solar' in lname:
                    gen_type = '태양광'
                elif 'wt' in lname or 'wind' in lname:
                    gen_type = '풍력'
                elif 'hydro' in lname or 'water' in lname:
                    gen_type = '수력'
                elif 'coal' in lname:
                    gen_type = '석탄'
                elif 'lng' in lname or 'gas' in lname:
                    gen_type = 'LNG'
                elif 'oil' in lname or 'diesel' in lname:
                    gen_type = '석유'
                elif 'biomass' in lname or 'bio' in lname:
                    gen_type = '바이오'
                elif 'geothermal' in lname:
                    gen_type = '지열'
                elif 'chp' in lname:
                    gen_type = '열병합'
                elif 'h2' in lname or 'hydrogen' in lname:
                    gen_type = '수소'
                elif '_h_fallback_gen' in lname or 'heat' in lname:
                    gen_type = '열'
                
                key = f"{region}_{gen_type}"
                if key not in regional_type_generation:
                    regional_type_generation[key] = 0
                regional_type_generation[key] += gen_output[gen_name].sum()
        
        # 지역별 발전원별 DataFrame 생성
        regional_type_data = []
        for key, value in regional_type_generation.items():
            region, gen_type = key.split('_', 1)
            regional_type_data.append([region, gen_type, value])
        
        regional_type_df = pd.DataFrame(regional_type_data, 
                                      columns=['지역', '발전원', '총발전량(MWh)'])
        regional_type_df.to_csv(f'{results_dir}/optimization_result_{current_time}_지역별_발전원별_발전량.csv', 
                               index=False, encoding='utf-8-sig')
        
        # 상위 발전기 발전량 (상위 20개)
        total_gen_output = gen_output.sum().sort_values(ascending=False)
        top_generators = total_gen_output.head(20)
        top_gen_df = pd.DataFrame({
            '발전기명': top_generators.index,
            '총발전량(MWh)': top_generators.values
        })
        top_gen_df.to_csv(f'{results_dir}/optimization_result_{current_time}_상위발전기_발전량.csv', 
                         index=False, encoding='utf-8-sig')
        
        print("지역별 분석 결과 생성 완료:")
        print(f"- 지역별 발전량")
        print(f"- 발전원별 발전량")
        print(f"- 지역별 발전원별 발전량")
        print(f"- 상위 발전기 발전량")
        
        return True
        
    except Exception as e:
        print(f"지역별 분석 중 오류 발생: {str(e)}")
        traceback.print_exc()
        return False

def ensure_column(df, column_name, default_value=np.nan):
    """DataFrame에 컬럼이 없으면 기본값으로 추가"""
    if column_name not in df.columns:
        df[column_name] = default_value
    return df


def apply_year_overrides(input_data, overrides):
    """연도별 오버라이드 적용. overrides 형식 예:
    {
        'generators': {
            'SEL_PV_1': {'p_nom': 100, 'p_nom_extendable': True},
            '*': {'marginal_cost': 0}
        },
        'timeseries': {'start_time': '2021-01-01', 'end_time': '2022-01-01'}
    }
    """
    if not overrides:
        return input_data
    data = copy.deepcopy(input_data)

    for sheet, spec in overrides.items():
        if sheet not in data:
            continue
        df = data[sheet]

        # timeseries는 단일 행을 가정
        if sheet == 'timeseries' and isinstance(spec, dict):
            if len(df) == 0:
                # 빈 경우 1행 생성
                data[sheet] = pd.DataFrame([spec])
            else:
                for col, val in spec.items():
                    if col not in df.columns:
                        df[col] = np.nan
                    df.at[df.index[0], col] = val
            continue

        # 그 외 시트: 'name' 기준으로 개별 행 적용 + '*' 와일드카드
        if isinstance(spec, dict):
            # 와일드카드 먼저
            wildcard_updates = spec.get('*', None)
            if wildcard_updates:
                for col, val in wildcard_updates.items():
                    if col not in df.columns:
                        df[col] = np.nan
                    df[col] = val

            # 개별 name 처리
            if 'name' in df.columns:
                for name_key, updates in spec.items():
                    if name_key == '*':
                        continue
                    if not isinstance(updates, dict):
                        continue
                    mask = df['name'].astype(str) == str(name_key)
                    if not mask.any():
                        continue
                    for col, val in updates.items():
                        if col not in df.columns:
                            df[col] = np.nan
                        df.loc[mask, col] = val

            data[sheet] = df

    return data


def extract_capacity_carryover(network):
    """이전 해 네트워크에서 다음 해로 인계할 용량 추출"""
    carry = {
        'generators': {},
        'lines': {},
        'links': {},
        'stores': {}
    }

    if network is None:
        return carry

    # Generators
    if not network.generators.empty:
        if 'p_nom_opt' in network.generators.columns:
            caps = network.generators['p_nom_opt'].fillna(network.generators['p_nom'])
        else:
            # 일부 버전은 optimize 후 p_nom_opt가 별도 속성으로 있음
            caps = getattr(network.generators, 'p_nom_opt', network.generators['p_nom'])
        for name, val in caps.items():
            carry['generators'][str(name)] = float(val)

    # Lines (AC)
    if not network.lines.empty:
        s_nom_opt = getattr(network.lines, 's_nom_opt', None)
        if s_nom_opt is not None:
            caps = network.lines['s_nom_opt'].fillna(network.lines['s_nom'])
        else:
            caps = network.lines['s_nom']
        for name, val in caps.items():
            carry['lines'][str(name)] = float(val)

    # Links (HVDC 등)
    if not network.links.empty:
        if 'p_nom_opt' in network.links.columns:
            caps = network.links['p_nom_opt'].fillna(network.links['p_nom'])
        else:
            caps = getattr(network.links, 'p_nom_opt', network.links['p_nom'])
        for name, val in caps.items():
            carry['links'][str(name)] = float(val)

    # Stores (energy capacity)
    if not network.stores.empty:
        if 'e_nom_opt' in network.stores.columns:
            caps = network.stores['e_nom_opt'].fillna(network.stores['e_nom'])
        else:
            caps = getattr(network.stores, 'e_nom_opt', network.stores['e_nom'])
        for name, val in caps.items():
            carry['stores'][str(name)] = float(val)

    return carry


def apply_carryover_to_input(input_data, carryover_caps, policy='min'):
    """인계 용량을 입력 데이터에 반영
    policy='min': 다음 해 최소 용량 하한으로 적용(추가 확장 허용)
    """
    if not carryover_caps:
        return input_data

    data = copy.deepcopy(input_data)

    # Generators
    if 'generators' in data and not data['generators'].empty:
        df = data['generators']
        ensure_column(df, 'p_nom_min', 0.0)
        for name, cap in carryover_caps.get('generators', {}).items():
            mask = df['name'].astype(str) == str(name)
            if mask.any():
                df.loc[mask, 'p_nom_min'] = np.maximum(df.loc[mask, 'p_nom_min'].astype(float), cap)
        data['generators'] = df

    # Lines
    if 'lines' in data and not data['lines'].empty:
        df = data['lines']
        ensure_column(df, 's_nom_min', 0.0)
        for name, cap in carryover_caps.get('lines', {}).items():
            mask = df['name'].astype(str) == str(name)
            if mask.any():
                df.loc[mask, 's_nom_min'] = np.maximum(df.loc[mask, 's_nom_min'].astype(float), cap)
        data['lines'] = df

    # Links
    if 'links' in data and not data['links'].empty:
        df = data['links']
        ensure_column(df, 'p_nom_min', 0.0)
        for name, cap in carryover_caps.get('links', {}).items():
            mask = df['name'].astype(str) == str(name)
            if mask.any():
                df.loc[mask, 'p_nom_min'] = np.maximum(df.loc[mask, 'p_nom_min'].astype(float), cap)
        data['links'] = df

    # Stores
    if 'stores' in data and not data['stores'].empty:
        df = data['stores']
        ensure_column(df, 'e_nom_min', 0.0)
        for name, cap in carryover_caps.get('stores', {}).items():
            mask = df['name'].astype(str) == str(name)
            if mask.any():
                df.loc[mask, 'e_nom_min'] = np.maximum(df.loc[mask, 'e_nom_min'].astype(float), cap)
        data['stores'] = df

    return data


def run_multi_year_sequence(years, base_input_file=INPUT_FILE, overrides_by_year=None, carryover=True, results_root='results_multi'):
    """연도별 순차 실행 루프.
    - years: [2020, 2021, ...]
    - overrides_by_year: {year: overrides(dict)}
    - carryover: True면 이전 해 용량을 다음 해 최소 용량으로 인계
    - results_root: 결과 저장 루트 디렉터리
    반환: {year: {'network': Network, 'results_dir': str}}
    """
    timestamp_root = os.path.join(results_root, datetime.now().strftime("%Y%m%d_%H%M%S"))
    os.makedirs(timestamp_root, exist_ok=True)
    results = {}
    prev_network = None

    for year in years:
        print(f"\n===== {year}년도 분석 시작 =====")
        # 0) 연도별 interface 시나리오를 통합 파일에 반영(시트 직접 갱신)
        try:
            root_dir = os.path.dirname(__file__)
            integrated_path = os.path.abspath(os.path.join(root_dir, base_input_file))
            interface_path = os.path.abspath(os.path.join(root_dir, 'interface.xlsx'))
            _update_integrated_for_year(integrated_path, interface_path, year)
        except Exception as _e:
            print(f"통합파일 갱신 경고({year}): {str(_e)}")
        # 1) 기본 입력 로드
        input_data = read_input_data(base_input_file)

        # 1.5) 해당 연도 수요 시나리오를 loads.p_set에 주입
        input_data = _apply_scenario_to_loads_in_input(input_data, year)

        # 1.6) 버스명 표준화 및 전 시트 반영(예: BSN_BSN_EL → BSN_EL)
        input_data = standardize_bus_names_in_input(input_data)
        try:
            integrated_path = os.path.abspath(os.path.join(os.path.dirname(__file__), base_input_file))
            _persist_standardized_input(integrated_path, input_data)
        except Exception as _e:
            print(f"표준화된 버스명 저장 경고: {str(_e)}")

        # 2) 연도별 오버라이드 적용 (수요/패턴/원가/효율/확장가능 등)
        if overrides_by_year and year in overrides_by_year:
            input_data = apply_year_overrides(input_data, overrides_by_year[year])

        # 3) 이전 해 결과 인계 (용량 하한)
        if carryover and prev_network is not None:
            caps = extract_capacity_carryover(prev_network)
            input_data = apply_carryover_to_input(input_data, caps, policy='min')

        # 4) 네트워크 생성 및 최적화
        network = create_network(input_data)
        success = optimize_network(network)
        if not success:
            print(f"{year}년도 최적화 실패. 부분 결과(시계열)를 저장합니다.")
            year_dir = os.path.join(timestamp_root, str(year))
            os.makedirs(year_dir, exist_ok=True)
            try:
                save_results(network, subdir=year_dir)
                results[year] = {'network': network, 'results_dir': year_dir}
            except Exception as _e_sv:
                print(f"부분 결과 저장 실패: {_e_sv}")
                results[year] = {'network': network, 'results_dir': None}
            prev_network = network
            continue

        # 5) 결과 저장 (타임스탬프/연도 서브폴더)
        year_dir = os.path.join(timestamp_root, str(year))
        os.makedirs(year_dir, exist_ok=True)
        save_results(network, subdir=year_dir)

        results[year] = {'network': network, 'results_dir': year_dir}
        prev_network = network
        print(f"===== {year}년도 분석 완료 =====\n")

    return results

def ensure_integrated_input():
    """interface.xlsx 기반으로 integrated_input_data.xlsx 생성/업데이트"""
    try:
        root_dir = os.path.dirname(__file__)
        integrated_path = os.path.abspath(os.path.join(root_dir, INPUT_FILE))
        interface_path = os.path.abspath(os.path.join(root_dir, 'interface.xlsx'))

        need_build = False
        reason = ''
        if not os.path.exists(integrated_path):
            need_build = True
            reason = '통합 파일이 존재하지 않음'
        elif os.path.exists(interface_path):
            try:
                if os.path.getmtime(interface_path) > os.path.getmtime(integrated_path):
                    need_build = True
                    reason = 'interface.xlsx가 더 최신'
            except Exception:
                pass

        if need_build:
            print(f"통합 입력 생성/업데이트 실행 ({reason})...")
            ok = False
            # 지역 엑셀 프로세서를 우선 사용(링크/CHP 포함 컬럼 처리 보장)
            try:
                from process_regional_excel import RegionalExcelProcessor
                import openpyxl as _oxl
                proc = RegionalExcelProcessor(interface_path)
                proc.output_path = integrated_path
                # 워크북을 열고 지역/링크 등 데이터를 먼저 적재한 뒤 통합 생성
                try:
                    _wb = _oxl.load_workbook(interface_path, data_only=True)
                    _steps_ok = True
                    _steps_ok &= bool(proc.read_selected_regions(_wb))
                    _steps_ok &= bool(proc.read_regional_data(_wb))
                    _steps_ok &= bool(proc.read_connections(_wb))
                    _steps_ok &= bool(proc.read_timeseries(_wb))
                    _steps_ok &= bool(proc.read_patterns(_wb))
                except Exception as _e0:
                    print(f"RegionalExcelProcessor 사전 적재 실패: {_e0}")
                    _steps_ok = False
                ok = bool(proc.create_integrated_data()) if _steps_ok else False
            except Exception as e:
                print(f"RegionalExcelProcessor 생성 실패: {str(e)}")
            # 보조 경로로 v3 생성기 시도
            if not ok:
                try:
                    from create_integrated_data import create_integrated_data as v3_create
                    ok = bool(v3_create())
                except Exception as e:
                    print(f"v3 생성기 호출 실패: {str(e)}")
            if ok:
                try:
                    print(f"통합 입력 생성 완료: {integrated_path}")
                    print(f"갱신된 수정 시간: {datetime.fromtimestamp(os.path.getmtime(integrated_path))}")
                    # 생성물 검증: 필수 시트 비어있으면 대체 생성기로 재시도
                    try:
                        _xls_chk = pd.ExcelFile(integrated_path)
                        _has_buses = ('buses' in _xls_chk.sheet_names) and (not pd.read_excel(integrated_path, sheet_name='buses').empty)
                        _has_gens = ('generators' in _xls_chk.sheet_names) and (not pd.read_excel(integrated_path, sheet_name='generators').empty)
                        if not (_has_buses and _has_gens):
                            print("경고: 통합 입력에 필수 시트가 비어있습니다. v3 생성기로 재시도합니다.")
                            from create_integrated_data import create_integrated_data as _v3_create
                            if bool(_v3_create()):
                                print("v3 생성기 재시도 성공")
                            else:
                                print("경고: v3 생성기 재시도 실패. 기존 파일 유지")
                    except Exception as _e_chk:
                        print(f"통합 입력 검증 실패: {_e_chk}")
                except Exception:
                    pass
            else:
                print("경고: 통합 입력 생성에 실패했습니다. 기존 파일이 있으면 그걸 사용합니다.")
        else:
            print("통합 입력이 최신 상태입니다. 재생성 생략.")
    except Exception as e:
        print(f"통합 입력 준비 중 오류: {str(e)}")

def _fallback_build_lines_from_interface(input_data, interface_path):
    try:
        if not os.path.exists(interface_path):
            print("interface.xlsx가 없어 선로 폴백 생성을 건너뜁니다.")
            return None
        # 지역간 연결 시트 읽기 (레이블은 5행, 데이터는 6행부터라는 전제)
        try:
            try:
                df = pd.read_excel(interface_path, sheet_name='지역간 연결', header=4)
            except Exception:
                # 폴백: 레이블이 1행일 수도 있으므로 0행 시도
                df = pd.read_excel(interface_path, sheet_name='지역간 연결', header=0)
        except Exception as e:
            print(f"'지역간 연결' 시트를 읽을 수 없습니다: {str(e)}")
            return None
        if df is None or df.empty:
            print("'지역간 연결' 시트가 비어있습니다.")
            return None

        # 컬럼 표준화(공백 제거) 및 로그
        df.columns = [str(c).strip() for c in df.columns]
        print(f"'지역간 연결' 시트 컬럼: {df.columns.tolist()}")
        print("상위 5행:\n", df.head(5))
        cols_lower = {c: str(c).strip().lower() for c in df.columns}

        def find_col(cands):
            for c in df.columns:
                cl = cols_lower[c]
                for k in cands:
                    if k == cl or k in cl:
                        return c
            return None

        # 주요 컬럼 매핑(레이블 매칭 우선)
        from_col = find_col(['from', '출발', '시작', '지역1', '출발지역', 'from_region', 'region1'])
        to_col = find_col(['to', '도착', '끝', '지역2', '도착지역', 'to_region', 'region2'])
        name_col = find_col(['name', '이름', '선로명'])

        # 사용자가 통일한 레이블 매칭(type, length, x, r, num_parallel)
        type_col = find_col(['type'])
        length_col = find_col(['length'])
        x_col = find_col(['x'])
        r_col = find_col(['r'])
        num_parallel_col = find_col(['num_parallel'])

        # 용량은 기존 키워드로 탐색
        cap_col = find_col(['s_nom', 'capacity', '용량', 'mva', '정격'])

        if from_col is None or to_col is None:
            print("'지역간 연결' 시트에서 출발/도착 컬럼을 찾지 못했습니다.")
            return None

        # 버스 목록
        buses_df = input_data.get('buses', pd.DataFrame())
        bus_names = list(buses_df['name'].astype(str)) if ('name' in buses_df.columns) else []
        bus_carriers = dict(zip(buses_df['name'].astype(str), buses_df['carrier'].astype(str))) if ('carrier' in buses_df.columns and 'name' in buses_df.columns) else {}

        def guess_bus(region_code):
            if not bus_names:
                return str(region_code)
            region_code = str(region_code).strip()
            candidates = [b for b in bus_names if b == region_code or b.startswith(region_code + '_')]
            if not candidates:
                # 넓게 포함 검색
                candidates = [b for b in bus_names if region_code in b]
            if not candidates:
                return region_code
            # 전력 계통 우선
            elec_pref = [b for b in candidates if ('electric' in b.lower()) or (bus_carriers.get(b, '').lower() == 'electricity')]
            if elec_pref:
                return elec_pref[0]
            return candidates[0]

        lines_rows = []
        added = 0
        for _, row in df.iterrows():
            fr_raw = row.get(from_col)
            to_raw = row.get(to_col)
            if pd.isna(fr_raw) or pd.isna(to_raw):
                continue
            fr_region = str(fr_raw).strip()
            to_region = str(to_raw).strip()

            name_val = str(row.get(name_col)).strip() if (name_col and pd.notna(row.get(name_col))) else f"{fr_region}-{to_region}"
            s_nom_val = float(pd.to_numeric(row.get(cap_col), errors='coerce')) if cap_col else 1000.0
            x_val = float(pd.to_numeric(row.get(x_col), errors='coerce')) if x_col else 0.1
            r_val = float(pd.to_numeric(row.get(r_col), errors='coerce')) if r_col else 0.01
            length_val = float(pd.to_numeric(row.get(length_col), errors='coerce')) if length_col else np.nan
            type_val = str(row.get(type_col)).strip() if (type_col and pd.notna(row.get(type_col))) else np.nan
            num_parallel_val = int(pd.to_numeric(row.get(num_parallel_col), errors='coerce')) if num_parallel_col else np.nan

            bus0 = guess_bus(fr_region)
            bus1 = guess_bus(to_region)

            lines_rows.append({
                'name': name_val,
                'bus0': bus0,
                'bus1': bus1,
                's_nom': s_nom_val,
                'x': x_val,
                'r': r_val,
                'length': length_val,
                'type': type_val,
                'num_parallel': num_parallel_val
            })
            added += 1

        if added == 0:
            print("'지역간 연결' 시트에서 유효한 선로 레코드를 만들지 못했습니다.")
            return None

        lines_df = pd.DataFrame(lines_rows)
        print(f"interface.xlsx → '지역간 연결'에서 선로 {len(lines_df)}개를 폴백 생성했습니다.")
        # 일부 기본 컬럼 보장
        for c in ['s_nom_extendable', 's_nom_min', 's_nom_max']:
            if c not in lines_df.columns:
                lines_df[c] = np.nan
        return lines_df
    except Exception as e:
        print(f"선로 폴백 생성 중 오류: {str(e)}")
        return None

def _persist_lines_to_integrated(integrated_path, lines_df):
    try:
        if lines_df is None or lines_df.empty:
            return False
        if not os.path.exists(integrated_path):
            print(f"경고: {integrated_path} 파일이 없어 lines를 저장하지 못했습니다.")
            return False
        # 기존 모든 시트를 읽어 dict로 보관 후 lines 교체/추가
        xls = pd.ExcelFile(integrated_path)
        sheets = {}
        for sn in xls.sheet_names:
            if sn.lower() == 'lines':
                continue
            sheets[sn] = pd.read_excel(integrated_path, sheet_name=sn)
        # 재작성
        with pd.ExcelWriter(integrated_path, engine='openpyxl') as writer:
            # 기존 시트들 쓰기
            for sn, df in sheets.items():
                df.to_excel(writer, sheet_name=sn, index=False)
            # lines 쓰기
            lines_df.to_excel(writer, sheet_name='lines', index=False)
        print(f"lines 시트를 '{integrated_path}'에 저장했습니다.")
        return True
    except Exception as e:
        print(f"lines 시트 저장 실패: {str(e)}")
        return False

def _to_bool(value):
    try:
        if isinstance(value, (bool, np.bool_)):
            return bool(value)
        if pd.isna(value):
            return False
        s = str(value).strip().lower()
        return s in ['1', 'true', 'yes', 'y', 't']
    except Exception:
        return False

def _normalize_bus_name(raw_name, available_bus_names, prefer_electric=True, bus_carriers=None):
    try:
        if not isinstance(raw_name, str):
            raw = str(raw_name)
        else:
            raw = raw_name.strip()
        bus_set = set(str(b) for b in available_bus_names)
        if raw in bus_set:
            return raw
        # 토큰 분해
        tokens = raw.split('_')
        tokens = [t for t in tokens if t]
        region = None
        energy = None
        if len(tokens) >= 2:
            region = tokens[0]
            energy = tokens[-1]
        elif len(tokens) == 1:
            region = tokens[0]
        candidates = []
        # 1) 원문
        candidates.append(raw)
        # 2) REGION_ENERGY 형태
        if region and energy:
            candidates.append(f"{region}_{energy}")
        # 3) REGION_REGION_ENERGY 형태
        if region and energy:
            candidates.append(f"{region}_{region}_{energy}")
        # 4) 시작/끝 패턴 매칭
        if region and energy:
            for b in available_bus_names:
                bs = str(b)
                if bs.startswith(region + '_') and bs.endswith('_' + energy):
                    candidates.append(bs)
        # 5) region로만 시작하는 모든 버스
        if region:
            for b in available_bus_names:
                bs = str(b)
                if bs.startswith(region + '_'):
                    candidates.append(bs)
        # 고유화
        seen = set()
        ordered = []
        for c in candidates:
            if c not in seen:
                seen.add(c)
                ordered.append(c)
        # 캐리어 선호
        if prefer_electric and bus_carriers:
            elec = [c for c in ordered if c in bus_set and str(bus_carriers.get(c, '')).lower() == 'electricity']
            if elec:
                return elec[0]
        # 첫 매칭 반환
        for c in ordered:
            if c in bus_set:
                return c
        return raw
    except Exception:
        return str(raw_name)

def _standardize_bus_token_by_carrier(carrier):
    c = str(carrier).strip().lower()
    if 'electric' in c or c == 'el' or c == '전력':
        return 'EL'
    if 'heat' in c or c == 'h' or c == '열':
        return 'H'
    if 'hydrogen' in c or c == 'h2' or c == '수소':
        return 'H2'
    if 'gas' in c or 'lng' in c or c == '가스':
        return 'LNG'
    return None


def standardize_bus_names_in_input(input_data):
    try:
        if 'buses' not in input_data or input_data['buses'].empty:
            return input_data
        buses = input_data['buses']
        if 'name' not in buses.columns:
            return input_data
        # 에너지원 토큰 결정
        energy_token_by_bus = {}
        if 'carrier' in buses.columns:
            for _, row in buses.iterrows():
                bus_name = str(row['name']).strip()
                token = _standardize_bus_token_by_carrier(row['carrier'])
                energy_token_by_bus[bus_name] = token
        
        # 표준 이름 생성 함수
        def make_std_name(old_name):
            name = str(old_name).strip()
            tokens = [t for t in name.split('_') if t]
            region = tokens[0] if tokens else name
            # 우선 carrier 기반 토큰 사용, 없으면 마지막 토큰 유지(LNG 허용)
            energy = energy_token_by_bus.get(name)
            if energy is None:
                last = tokens[-1] if len(tokens) >= 2 else 'EL'
                energy = last if last.upper() in ['EL','H','H2','LNG'] else 'EL'
            return f"{region}_{energy}"
                
        # 매핑 생성
        mapping = {}
        for _, row in buses.iterrows():
            old = str(row['name']).strip()
            new = make_std_name(old)
            mapping[old] = new
        
        # 충돌 방지: 동일 new로 여러 old 매핑되는 경우 원본 유지
        reverse = {}
        for old, new in mapping.items():
            if new not in reverse:
                reverse[new] = old
            else:
                # 충돌 시 변경하지 않음
                mapping[old] = old
        
        # 적용 함수
        def apply_map(df, col):
            if df is None or df.empty or col not in df.columns:
                return df
            df[col] = df[col].astype(str).map(lambda x: mapping.get(x, x))
            return df
        
        # buses.name 업데이트
        input_data['buses']['name'] = input_data['buses']['name'].astype(str).map(lambda x: mapping.get(x, x))
        # generators.bus, loads.bus, links bus0/bus1, lines bus0/bus1 업데이트
        if 'generators' in input_data and not input_data['generators'].empty:
            input_data['generators'] = apply_map(input_data['generators'], 'bus')
        if 'loads' in input_data and not input_data['loads'].empty:
            input_data['loads'] = apply_map(input_data['loads'], 'bus')
        if 'links' in input_data and not input_data['links'].empty:
            input_data['links'] = apply_map(input_data['links'], 'bus0')
            input_data['links'] = apply_map(input_data['links'], 'bus1')
            input_data['links'] = apply_map(input_data['links'], 'bus2')
            input_data['links'] = apply_map(input_data['links'], 'bus3')
        if 'lines' in input_data and not input_data['lines'].empty:
            input_data['lines'] = apply_map(input_data['lines'], 'bus0')
            input_data['lines'] = apply_map(input_data['lines'], 'bus1')
        if 'stores' in input_data and not input_data['stores'].empty:
            input_data['stores'] = apply_map(input_data['stores'], 'bus')
        
        # 로그 요약
        changes = sum(1 for old, new in mapping.items() if old != new)
        if changes > 0:
            print(f"버스명 표준화 적용: {changes}개 이름 변경(형식 REGION_ENERGY, 예: BSN_EL)")

        # 추가 정규화: lines/link/store의 버스 컬럼에 잔여 중복 토큰(예: BSN_EL_EL) 압축 및 실제 버스 세트에 맞게 보정
        try:
            bus_names_set = set(input_data['buses']['name'].astype(str)) if ('buses' in input_data and not input_data['buses'].empty and 'name' in input_data['buses'].columns) else set()

            def _compress_bus_name(raw):
                try:
                    s = str(raw).strip()
                    if not s:
                        return s
                    if s in bus_names_set:
                        return s
                    tokens = [t for t in s.split('_') if t]
                    region = tokens[0] if tokens else s
                    # 에너지 토큰 후보 압축 (중복 제거) + LNG 허용
                    energy_candidates = [t for t in tokens[1:] if t.upper() in ['EL','H','H2','LNG']]
                    energy = None
                    for t in energy_candidates:
                        tt = t.upper()
                        if tt in ['EL','H','H2','LNG']:
                            energy = tt
                            break
                    if energy is None and len(tokens) >= 2:
                        energy = tokens[-1].upper()
                        if energy not in ['EL','H','H2','LNG']:
                            # 기본값은 EL로 폴백(단, 원문에 LNG가 포함되어 있으면 LNG 유지)
                            if 'LNG' in [t.upper() for t in tokens[1:]]:
                                energy = 'LNG'
                            else:
                                energy = 'EL'
                    candidate = f"{region}_{energy}" if energy else region
                    if candidate in bus_names_set:
                        return candidate
                    # region으로 시작하는 버스 중 첫번째 대안
                    for b in bus_names_set:
                        if b.startswith(region + '_'):
                            return b
                    return candidate
                except Exception:
                    return str(raw)

            def _apply_compress(df, columns):
                if df is None or df.empty:
                    return df
                for c in columns:
                    if c in df.columns:
                        df[c] = df[c].astype(str).map(_compress_bus_name)
                return df

            if 'lines' in input_data and not input_data['lines'].empty:
                input_data['lines'] = _apply_compress(input_data['lines'], ['bus0','bus1'])
            if 'links' in input_data and not input_data['links'].empty:
                input_data['links'] = _apply_compress(input_data['links'], ['bus0','bus1','bus2','bus3'])
            if 'stores' in input_data and not input_data['stores'].empty:
                input_data['stores'] = _apply_compress(input_data['stores'], ['bus'])
        except Exception as _e:
            print(f"버스명 추가 정규화 경고: {str(_e)}")

        return input_data
    except Exception as e:
        print(f"버스명 표준화 중 오류: {str(e)}")
        return input_data

def _persist_standardized_input(integrated_path, input_data):
    try:
        if not input_data or not isinstance(input_data, dict):
            return False
        # 쓰기 가능한 시트만 수집
        sheets = {k: v for k, v in input_data.items() if isinstance(v, pd.DataFrame)}
        if not sheets:
            return False
        with pd.ExcelWriter(integrated_path, engine='openpyxl') as writer:
            for sn, df in sheets.items():
                # 인덱스 제거하여 저장
                df.to_excel(writer, sheet_name=sn, index=False)
        print(f"표준화된 버스명이 '{integrated_path}'에 저장되었습니다.")
        return True
    except Exception as e:
        print(f"표준화 입력 저장 실패: {str(e)}")
        return False

def _parse_generator_scenario_from_interface(interface_path, year):
    try:
        if not os.path.exists(interface_path):
            return {}
        df = pd.read_excel(interface_path, sheet_name='시나리오_발전기')
        if df is None or df.empty:
            return {}
        df.columns = [str(c).strip() for c in df.columns]

        # 연도 컬럼(예: '2024', '2025', ...) 탐지
        year_str = str(year)
        year_cols = [c for c in df.columns if str(c).strip().isdigit() and len(str(c).strip()) == 4]
        name_col = None
        bus_col = None
        for c in df.columns:
            cl = str(c).strip().lower()
            if cl in ['name', '이름']:
                name_col = c
            if cl in ['bus', '버스']:
                bus_col = c
        scenario_by_name = {}

        if year_cols and year_str in [str(c).strip() for c in year_cols] and name_col is not None:
            # 형식: [이름, 버스, 2024, 2025, ...]
            year_col = [c for c in df.columns if str(c).strip() == year_str][0]
            for _, row in df.iterrows():
                gname = str(row[name_col]).strip()
                if not gname or gname.lower() == 'nan':
                    continue
                val = pd.to_numeric(row.get(year_col), errors='coerce')
                if pd.notna(val):
                    scenario_by_name[gname] = float(val)
            return scenario_by_name

        # 폴백: 기존 방식(연도/지역/타입) 형식은 더 이상 사용하지 않지만, 남겨둠
        return {}
    except Exception as e:
        print(f"발전기 시나리오 파싱 오류: {str(e)}")
        return {}


def _build_generator_overrides_for_year_v2(input_data, year, interface_path):
    try:
        gens = input_data.get('generators', pd.DataFrame())
        if gens is None or gens.empty:
            return {}
        scenario = _parse_generator_scenario_from_interface(interface_path, year)
        if not scenario:
            print(f"{year}년 발전기 시나리오가 없어 오버라이드 생략")
            return {}
        def get_type_from_name(name):
            n = str(name)
            if 'PV' in n:
                return 'PV'
            if 'WT' in n:
                return 'WT'
            if 'Nuclear' in n:
                return 'Nuclear'
            if 'CHP' in n:
                return 'CHP'
            if 'Coal' in n:
                return 'Coal'
            if 'Hydro' in n:
                return 'Hydro'
            if 'LNG' in n:
                return 'LNG'
            return None
        # 그룹별(지역, 타입) 기본 합계 및 멤버 목록
        group_sum = {}
        group_members = {}
        for _, row in gens.iterrows():
            name = str(row['name'])
            region = name.split('_')[0] if '_' in name else None
            gtype = get_type_from_name(name)
            if region and gtype:
                key = (region, gtype)
                base_nom = float(pd.to_numeric(row['p_nom'], errors='coerce') or 0.0)
                group_sum[key] = group_sum.get(key, 0.0) + base_nom
                group_members.setdefault(key, []).append((name, base_nom))
        overrides = {'generators': {}}
        for key, target in scenario.items():
            region, gtype = key
            base = group_sum.get(key, 0.0)
            members = group_members.get(key, [])
            if not members:
                print(f"경고: {year}년 {region}-{gtype} 그룹에 대응하는 발전기 레코드가 없습니다.")
                continue
            if base <= 0:
                print(f"경고: {year}년 {region}-{gtype} 기본 합계 0 → 분배 불가, 기존 값 유지")
                continue
            scale = float(target) / float(base)
            for name, base_nom in members:
                new_nom = base_nom * scale
                if gtype in ['PV', 'WT']:
                    overrides['generators'][name] = {
                        'p_nom_min': new_nom,
                        'p_nom': new_nom
                    }
                else:
                    overrides['generators'][name] = {
                        'p_nom': new_nom
                    }
        return overrides
    except Exception as e:
        print(f"발전기 오버라이드 v2 생성 오류: {str(e)}")
        return {}


def build_overrides_for_years(years, base_input_file):
    try:
        root_dir = os.path.dirname(__file__)
        interface_path = os.path.abspath(os.path.join(root_dir, 'interface.xlsx'))
        base = read_input_data(base_input_file)
        freq = '1h'
        try:
            if 'timeseries' in base and not base['timeseries'].empty:
                freq = str(base['timeseries'].iloc[0]['frequency'])
        except Exception:
            pass
        overrides_by_year = {}
        for y in years:
            ts_override = {
                'start_time': f"{y}-01-01 00:00:00",
                'end_time': f"{y+1}-01-01 00:00:00",
                'frequency': freq
            }
                    # 이름별 목표용량을 가져와 개별 발전기에 직접 주입
        name_to_target = _parse_generator_scenario_from_interface(interface_path, y)
        ov = {'timeseries': ts_override, 'generators': {}}
        if name_to_target:
            for gname, target in name_to_target.items():
                # 재생 여부에 따라 최소용량만 지정(확장가능 여부는 입력 파일/인터페이스에 따름)
                if any(k in gname for k in ['PV','WT']):
                    ov['generators'][gname] = {'p_nom_min': float(target), 'p_nom': float(target)}
                else:
                    ov['generators'][gname] = {'p_nom': float(target)}
        overrides_by_year[y] = ov
        print(f"지정 연도 오버라이드 구성 완료: {list(overrides_by_year.keys())}")
        return overrides_by_year
    except Exception as e:
        print(f"지정 연도 오버라이드 구성 오류: {str(e)}")
        return {}

def _parse_demand_scenario_wide(interface_path):
    try:
        if not os.path.exists(interface_path):
            return {}
        df = pd.read_excel(interface_path, sheet_name='시나리오_에너지수요')
        if df is None or df.empty:
            return {}
        df.columns = [str(c).strip() for c in df.columns]
        name_col = None
        for c in df.columns:
            if str(c).strip().lower() in ['name', '이름']:
                name_col = c
                break
        if name_col is None:
            return {}
        year_cols = [c for c in df.columns if str(c).strip().isdigit() and len(str(c).strip()) == 4]
        if not year_cols:
            return {}
        out = {}
        for _, row in df.iterrows():
            lname = str(row.get(name_col, '')).strip()
            if not lname or lname.lower() == 'nan':
                continue
            per_year = {}
            for yc in year_cols:
                val = pd.to_numeric(row.get(yc), errors='coerce')
                if pd.notna(val):
                    per_year[int(str(yc).strip())] = float(val)
            if per_year:
                out[lname] = per_year
        return out
    except Exception as e:
        print(f"수요 시나리오(와이드) 파싱 오류: {str(e)}")
        return {}


def _apply_scenario_to_loads_in_input(input_data, scenario_year):
    try:
        if 'loads' not in input_data or input_data['loads'].empty:
            return input_data
        # interface.xlsx 경로
        root_dir = os.path.dirname(__file__)
        interface_path = os.path.abspath(os.path.join(root_dir, 'interface.xlsx'))
        if not os.path.exists(interface_path):
            return input_data
        try:
            df = pd.read_excel(interface_path, sheet_name='시나리오_에너지수요')
        except Exception:
            return input_data
        if df is None or df.empty:
            return input_data
        df.columns = [str(c).strip() for c in df.columns]
        year_col = str(scenario_year)
        if year_col not in df.columns:
            return input_data
        # 기준 컬럼 탐색
        name_col = None
        bus_col = None
        for c in df.columns:
            cl = str(c).strip().lower()
            if cl in ['name', '이름']:
                name_col = c
            if cl in ['bus', '버스']:
                bus_col = c
        loads_df = input_data['loads']
        if 'p_set' not in loads_df.columns:
            loads_df['p_set'] = np.nan
        # 시나리오 → 값 매핑 구성
        map_by_name = {}
        if name_col is not None:
            for _, row in df.iterrows():
                nm = str(row.get(name_col, '')).strip()
                if not nm:
                    continue
                val = pd.to_numeric(row.get(year_col), errors='coerce')
                if pd.notna(val):
                    map_by_name[nm] = float(val)
        # 버스 기반 보조 매핑(BSN_EL → BSN_Demand_EL 등)
        if bus_col is not None:
            for _, row in df.iterrows():
                bus = str(row.get(bus_col, '')).strip()
                if not bus:
                    continue
                val = pd.to_numeric(row.get(year_col), errors='coerce')
                if pd.isna(val):
                    continue
                parts = [t for t in bus.split('_') if t]
                if len(parts) >= 2:
                    region, energy = parts[0], parts[-1]
                    load_name = f"{region}_Demand_{energy}"
                    # 이름 매핑이 우선이나 없으면 버스 파생 이름으로 보강
                    if load_name not in map_by_name:
                        map_by_name[load_name] = float(val)
        # 매핑 적용(이름 기준 → 순서 불일치 해소)
        updated = 0
        for idx, row in loads_df.iterrows():
            lname = str(row.get('name', '')).strip()
            if lname in map_by_name:
                loads_df.at[idx, 'p_set'] = map_by_name[lname]
                updated += 1
        input_data['loads'] = loads_df
        print(f"loads.p_set 주입(이름/버스 매핑): {scenario_year}년 {updated}개 행 업데이트")
        if updated == 0:
            print("경고: 시나리오_에너지수요에서 일치하는 이름/버스를 찾지 못했습니다. 시트의 '이름' 또는 '버스' 컬럼과 loads의 'name'을 확인하세요.")
        return input_data
    except Exception as e:
        print(f"loads 시나리오 주입 오류: {str(e)}")
        return input_data

def _update_integrated_for_year(integrated_path, interface_path, year):
    try:
        if not os.path.exists(integrated_path) or not os.path.exists(interface_path):
            return False
        # 통합 파일의 모든 시트를 읽음
        xls_int = pd.ExcelFile(integrated_path)
        sheets = {sn: pd.read_excel(integrated_path, sheet_name=sn) for sn in xls_int.sheet_names}

        # 1) 시나리오_에너지수요 → loads.p_set 업데이트 (연도 헤더 기반, 순서 매칭)
        try:
            df_load_scn = pd.read_excel(interface_path, sheet_name='시나리오_에너지수요')
            df_load_scn.columns = [str(c).strip() for c in df_load_scn.columns]
            year_str = str(year)
            if 'loads' in sheets and year_str in df_load_scn.columns:
                df_loads = sheets['loads']
                # 연도 컬럼에서 숫자만 추출하여 상단/하단 NaN 제거
                series = pd.to_numeric(df_load_scn[year_str], errors='coerce')
                vals = series.dropna().values
                n = min(len(df_loads), len(vals))
                if n > 0:
                    df_loads.loc[df_loads.index[:n], 'p_set'] = vals[:n]
                    sheets['loads'] = df_loads
                    print(f"integrated_input_data loads.p_set 갱신: {year}년 {n}개 행 업데이트")
        except Exception as e:
            print(f"loads 갱신 경고({year}): {str(e)}")

        # 2) 시나리오_발전기 → generators 업데이트 (이름 기반)
        try:
            df_gen_scn = pd.read_excel(interface_path, sheet_name='시나리오_발전기')
            df_gen_scn.columns = [str(c).strip() for c in df_gen_scn.columns]
            name_col = None
            for c in df_gen_scn.columns:
                if str(c).strip().lower() in ['name', '이름']:
                    name_col = c
                    break
            year_str = str(year)
            if 'generators' in sheets and name_col is not None and year_str in df_gen_scn.columns:
                df_gens = sheets['generators']
                if 'p_nom_min' not in df_gens.columns:
                    df_gens['p_nom_min'] = np.nan
                updated = 0
                # 이름 → 값 매핑
                name_to_val = {}
                for _, row in df_gen_scn.iterrows():
                    gname = str(row.get(name_col, '')).strip()
                    if not gname:
                        continue
                    val = pd.to_numeric(row.get(year_str), errors='coerce')
                    if pd.notna(val):
                        name_to_val[gname] = float(val)
                for idx, row in df_gens.iterrows():
                    gname = str(row.get('name', '')).strip()
                    if not gname or gname not in name_to_val:
                        continue
                    target = name_to_val[gname]
                    if ('PV' in gname) or ('WT' in gname):
                        # 재생: 최소용량 설정 (확장가능 여부는 기존값/인터페이스에 따름)
                        df_gens.at[idx, 'p_nom_min'] = float(target)
                        # p_nom은 최소 target 이상으로 보정
                        base_nom = float(pd.to_numeric(row.get('p_nom'), errors='coerce') or 0.0)
                        df_gens.at[idx, 'p_nom'] = float(max(base_nom, target))
                    else:
                        # 비재생: 연도 값으로 고정
                        df_gens.at[idx, 'p_nom'] = float(target)
                    updated += 1
                sheets['generators'] = df_gens
                print(f"integrated_input_data generators 갱신: {year}년 {updated}개 행 업데이트")
        except Exception as e:
            print(f"generators 갱신 경고({year}): {str(e)}")

        # 3) 지역간 연결 → lines 재생성(연결/길이/형식/병렬수 포함)
        try:
            tmp_input = {'buses': sheets.get('buses', pd.DataFrame())}
            fb_lines = _fallback_build_lines_from_interface(tmp_input, interface_path)
            if fb_lines is not None and not fb_lines.empty:
                sheets['lines'] = fb_lines
                print(f"integrated_input_data lines 갱신: {len(fb_lines)}개 레코드")
        except Exception as e:
            print(f"lines 갱신 경고({year}): {str(e)}")

        # 변경사항 저장
        with pd.ExcelWriter(integrated_path, engine='openpyxl') as writer:
            for sn, df in sheets.items():
                df.to_excel(writer, sheet_name=sn, index=False)
        return True
    except Exception as e:
        print(f"연도별 통합 입력 갱신 실패: {str(e)}")
        return False

def main():
    # 지도 시각화 - 임시로 주석 처리
    # visualizer = KoreaMapVisualizer()
    # if visualizer.load_map_data():
    #     visualizer.plot_korea_map()
    
    # 통합 입력 자동 생성/업데이트
    ensure_integrated_input()

    # 시나리오 오버라이드 비활성화 - 지역별 시트에서 직접 값 사용
    print("지역별 시트 원본 데이터 사용 모드로 실행...")
    
    # 단일년도 폴백 실행 (시나리오 오버라이드 없이)
    print("데이터 로드 시작...")
    input_data = read_input_data(INPUT_FILE)
    if input_data is None:
        return
        
    # 버스명 표준화(예: BSN_BSN_EL → BSN_EL)
    input_data = standardize_bus_names_in_input(input_data)

    # 표준화된 입력을 통합 파일에 즉시 저장
    try:
        integrated_path = os.path.abspath(os.path.join(os.path.dirname(__file__), INPUT_FILE))
        _persist_standardized_input(integrated_path, input_data)
    except Exception as e:
        print(f"표준화 저장 중 오류: {str(e)}")
        
    check_excel_data_loading(input_data)  # 데이터 로드 상태 확인
    
    print("네트워크 생성 시작...")
    network = create_network(input_data)
    
    print("최적화 시작...")
    if optimize_network(network):
        print("결과 저장 시작...")
        save_results(network)
        print("모든 과정 완료!")
    else:
        print("최적화 실패!")

if __name__ == "__main__":
    main()