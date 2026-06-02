"""
PyPSA-GESI 장기 시나리오 순차 분석
====================================
분석 연도 : 2030 → 2035 → 2040 → 2045 → 2050  (5년 간격)

핵심 연결 구조
--------------
각 연도의 최적화 결과(p_nom_opt, s_nom_opt, e_nom_opt)가
다음 연도의 최소 설비 하한(p_nom_min, s_nom_min, e_nom_min)으로 자동 반영됩니다.

예시)  2030년 태양광 SEL_PV_1 → p_nom_opt = 800 MW
       2035년 입력에서 SEL_PV_1 → p_nom_min = 800 MW (이 이하로 줄일 수 없음)
       2035년 p_nom_extendable=True 이면 800 MW 이상으로 추가 확장 가능

각 연도 분석 흐름
-----------------
1. interface.xlsx → integrated_input_data.xlsx 동기화
2. 기본 입력 로드  (integrated_input_data.xlsx)
3. 해당 연도 수요 시나리오 주입  (시나리오_에너지수요 시트)
4. 버스명 표준화
5. 연도별 오버라이드 적용  (timeseries 기간, 발전기 시나리오)
6. 이전 연도 설비 용량 인계  (p_nom_min 설정)
7. PyPSA 네트워크 생성
8. 최적화 실행
9. 결과 저장 (Excel / CSV / NetCDF / 시각화)
10. 설비 용량 스냅샷 저장  (results_longterm/_capacity_snapshots/capacity_YYYY.json)

실행 방법
---------
  # 전체 순차 실행 (2030 → 2035 → … → 2050)
  python run_longterm_scenario.py

  # 특정 연도만 실행
  python run_longterm_scenario.py --step 2030

  # 저장된 2030 스냅샷을 사용해 2035부터 재개
  python run_longterm_scenario.py --start 2035

  # 원하는 연도 조합만 실행
  python run_longterm_scenario.py --years 2030 2035

  # 결과 폴더 지정
  python run_longterm_scenario.py --results-root my_results
"""

import os
import sys
import argparse

# PyPSA_GUI.py 와 같은 폴더에서 임포트
_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _ROOT)

import PyPSA_GUI as _gui

# 장기 시나리오 분석은 반복 실행이 많으므로 시각화·후처리를 자동 생략
_gui.SKIP_VISUALIZATION = True

from PyPSA_GUI import (
    INPUT_FILE,
    read_input_data,
    create_network,
    optimize_network,
    save_results,
    extract_capacity_carryover,
    apply_carryover_to_input,
    apply_year_overrides,
    standardize_bus_names_in_input,
    build_overrides_for_years,
    save_capacity_snapshot,
    load_capacity_snapshot,
    ensure_integrated_input,
    _apply_scenario_to_loads_in_input,
    _update_integrated_for_year,
    _persist_standardized_input,
)

# ─── 전역 설정 ────────────────────────────────────────────────────────────────
LONGTERM_YEARS = [2030, 2035, 2040, 2045, 2050]
RESULTS_ROOT   = 'results_longterm'
SNAPSHOTS_DIR  = os.path.join(RESULTS_ROOT, '_capacity_snapshots')


# ─── 단일 연도 분석 ───────────────────────────────────────────────────────────
def run_one_year(year, prev_carry, overrides_by_year,
                 results_root=RESULTS_ROOT,
                 base_input_file=INPUT_FILE):
    """
    단일 연도 분석 실행.

    Parameters
    ----------
    year             : 분석 대상 연도 (int, 예: 2030)
    prev_carry       : 이전 연도의 설비 용량 dict  (없으면 None)
                       형식: {'generators': {name: MW}, 'lines': {name: MW},
                              'links': {name: MW}, 'stores': {name: MWh}}
    overrides_by_year: {year: overrides_dict}  (build_overrides_for_years 반환값)
    results_root     : 결과 저장 루트 폴더
    base_input_file  : 통합 입력 파일명

    Returns
    -------
    carry   : 이 연도의 최적화 설비 용량 dict  (다음 연도 run_one_year 에 전달)
    resdir  : 결과 저장 디렉터리 경로
    """
    print(f"\n{'='*62}")
    print(f"  장기 시나리오 분석: {year}년")
    print(f"{'='*62}")

    integrated = os.path.abspath(os.path.join(_ROOT, base_input_file))
    interface  = os.path.abspath(os.path.join(_ROOT, 'interface.xlsx'))
    year_dir   = os.path.join(results_root, str(year))
    os.makedirs(year_dir, exist_ok=True)

    # ── Step 0: interface.xlsx → integrated_input_data.xlsx 동기화 ─────────
    try:
        _update_integrated_for_year(integrated, interface, year)
    except Exception as e:
        print(f"[경고] 통합파일 갱신 오류 ({year}): {e}")

    # ── Step 1: 기본 입력 로드 ───────────────────────────────────────────────
    input_data = read_input_data(base_input_file)
    if input_data is None:
        print(f"[오류] {year}년 입력 데이터 로드 실패")
        return None, None

    # ── Step 2: 해당 연도 수요 시나리오 주입 ────────────────────────────────
    input_data = _apply_scenario_to_loads_in_input(input_data, year)

    # ── Step 3: 버스명 표준화 ───────────────────────────────────────────────
    input_data = standardize_bus_names_in_input(input_data)
    try:
        _persist_standardized_input(integrated, input_data)
    except Exception as e:
        print(f"[경고] 표준화된 버스명 저장 오류: {e}")

    # ── Step 4: 연도별 오버라이드 적용 ──────────────────────────────────────
    #   - timeseries 기간을 해당 연도로 변경
    #   - interface.xlsx '시나리오_발전기' 에 정의된 용량 반영
    if year in overrides_by_year:
        input_data = apply_year_overrides(input_data, overrides_by_year[year])
        print(f"[{year}] 연도별 오버라이드 적용 완료")

    # ── Step 5: 이전 연도 설비 용량 인계 ────────────────────────────────────
    #   이전 연도 p_nom_opt → 현재 연도 p_nom_min (절대 축소 불가)
    #   p_nom_extendable=True 인 설비는 p_nom_min 이상으로 추가 확장 가능
    if prev_carry:
        prev_year = year - 5
        input_data = apply_carryover_to_input(input_data, prev_carry, policy='min')
        print(f"[{year}] {prev_year}년 설비 용량 인계 완료 → p_nom_min 설정")

        # 인계 현황 요약 출력
        n_gen   = len(prev_carry.get('generators', {}))
        n_line  = len(prev_carry.get('lines', {}))
        n_link  = len(prev_carry.get('links', {}))
        n_store = len(prev_carry.get('stores', {}))
        print(f"       인계 항목: 발전기 {n_gen}개 / 선로 {n_line}개 / "
              f"링크 {n_link}개 / 저장장치 {n_store}개")

    # ── Step 6: PyPSA 네트워크 생성 ─────────────────────────────────────────
    print(f"\n[{year}] 네트워크 생성 중...")
    network = create_network(input_data)
    if network is None:
        print(f"[오류] {year}년 네트워크 생성 실패")
        return None, None

    # ── Step 7: 최적화 실행 ─────────────────────────────────────────────────
    print(f"[{year}] 최적화 실행 중...")
    success = optimize_network(network)

    if not success:
        print(f"[오류] {year}년 최적화 실패 → 부분 결과 저장 후 중단")
        try:
            save_results(network, subdir=year_dir)
        except Exception:
            pass
        return None, year_dir

    # ── Step 8: 결과 저장 ───────────────────────────────────────────────────
    print(f"[{year}] 결과 저장 중 → {year_dir}")
    save_results(network, subdir=year_dir)

    # ── Step 9: 설비 용량 스냅샷 저장 ──────────────────────────────────────
    carry = extract_capacity_carryover(network)
    save_capacity_snapshot(carry, year, snapshots_dir=SNAPSHOTS_DIR)

    print(f"\n[{year}년 분석 완료]  결과 폴더: {year_dir}")
    return carry, year_dir


# ─── 순차 분석 메인 루프 ─────────────────────────────────────────────────────
def run_longterm_sequence(years=None, start_from=None,
                          results_root=RESULTS_ROOT,
                          base_input_file=INPUT_FILE):
    """
    장기 시나리오 순차 분석 실행.

    Parameters
    ----------
    years       : 분석 연도 리스트  (기본: LONGTERM_YEARS)
    start_from  : 이 연도부터 시작  (이전 연도 스냅샷 파일이 필요)
    results_root: 결과 저장 루트 폴더
    base_input_file: 통합 입력 파일명

    Returns
    -------
    {year: {'carry': dict, 'results_dir': str}}
    """
    if years is None:
        years = LONGTERM_YEARS

    # ── resume 모드: 중간 연도부터 시작 ─────────────────────────────────────
    prev_carry = None
    if start_from is not None and start_from in years:
        idx = years.index(start_from)
        if idx > 0:
            prev_year = years[idx - 1]
            prev_carry = load_capacity_snapshot(prev_year, snapshots_dir=SNAPSHOTS_DIR)
            if prev_carry is None:
                print(f"\n[오류] {start_from}년부터 재개하려면 "
                      f"{prev_year}년 스냅샷이 필요합니다.")
                print(f"       먼저 {prev_year}년 분석을 실행하세요.")
                return {}
        years = years[idx:]

    # ── 연도별 오버라이드 구성 ────────────────────────────────────────────────
    print("\n연도별 오버라이드 구성 중...")
    overrides_by_year = build_overrides_for_years(years, base_input_file)

    # ── 통합 입력 초기화 ─────────────────────────────────────────────────────
    ensure_integrated_input()

    # ── 순차 실행 루프 ────────────────────────────────────────────────────────
    all_results = {}
    for year in years:
        carry, resdir = run_one_year(
            year=year,
            prev_carry=prev_carry,
            overrides_by_year=overrides_by_year,
            results_root=results_root,
            base_input_file=base_input_file,
        )
        all_results[year] = {'carry': carry, 'results_dir': resdir}
        prev_carry = carry   # ← 다음 연도에 전달하는 핵심 연결

        if carry is None:
            print(f"\n[중단] {year}년 분석 실패 → 이후 연도 분석을 중단합니다.")
            print(f"       원인 파악 후 --start {year} 옵션으로 재개 가능합니다.")
            break

    # ── 완료 요약 ─────────────────────────────────────────────────────────────
    done   = [y for y, v in all_results.items() if v['carry'] is not None]
    failed = [y for y, v in all_results.items() if v['carry'] is None]

    print(f"\n{'='*62}")
    print(f"  장기 시나리오 분석 완료")
    print(f"  성공: {done}")
    if failed:
        print(f"  실패: {failed}")
    print(f"  결과 폴더: {os.path.abspath(results_root)}")
    print(f"  스냅샷 폴더: {os.path.abspath(SNAPSHOTS_DIR)}")
    print(f"{'='*62}")
    return all_results


# ─── CLI 진입점 ───────────────────────────────────────────────────────────────
if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='PyPSA-GESI 장기 시나리오 순차 분석 (2030~2050, 5년 간격)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예시:
  python run_longterm_scenario.py                          # 전체 실행
  python run_longterm_scenario.py --step 2030             # 2030년만 실행
  python run_longterm_scenario.py --start 2035            # 2030 스냅샷 사용, 2035부터 재개
  python run_longterm_scenario.py --years 2030 2035       # 2030, 2035만 실행
  python run_longterm_scenario.py --results-root my_out   # 결과 폴더 지정
        """
    )
    parser.add_argument(
        '--years', nargs='+', type=int, default=None,
        metavar='YEAR',
        help=f'분석 연도 목록 (기본: {LONGTERM_YEARS})'
    )
    parser.add_argument(
        '--start', type=int, default=None,
        metavar='YEAR',
        help='이 연도부터 시작 (이전 연도 스냅샷 필요)'
    )
    parser.add_argument(
        '--step', type=int, default=None,
        metavar='YEAR',
        help='단일 연도만 실행'
    )
    parser.add_argument(
        '--results-root', type=str, default=RESULTS_ROOT,
        metavar='DIR',
        help=f'결과 저장 폴더 (기본: {RESULTS_ROOT})'
    )
    args = parser.parse_args()

    if args.step is not None:
        # ── 단일 연도 실행 ───────────────────────────────────────────────────
        target_year = args.step
        prev_carry  = None
        if target_year in LONGTERM_YEARS:
            idx = LONGTERM_YEARS.index(target_year)
            if idx > 0:
                prev_carry = load_capacity_snapshot(
                    LONGTERM_YEARS[idx - 1], snapshots_dir=SNAPSHOTS_DIR
                )
        overrides = build_overrides_for_years([target_year], INPUT_FILE)
        ensure_integrated_input()
        run_one_year(
            year=target_year,
            prev_carry=prev_carry,
            overrides_by_year=overrides,
            results_root=args.results_root,
        )
    else:
        # ── 순차(또는 부분) 실행 ────────────────────────────────────────────
        run_longterm_sequence(
            years=args.years,
            start_from=args.start,
            results_root=args.results_root,
        )
