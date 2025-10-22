#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
시나리오별 송전망 비교 시각화

사용법:
    python visualize_transmission_scenarios.py

입력 파일: scenario key results.xlsx (네트워크 혼잡 시트)
출력 폴더: results/scenario_comparison/
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import os
from matplotlib.animation import FuncAnimation, PillowWriter

# shapefile 읽기 라이브러리
try:
    import geopandas as gpd
    HAS_GEOPANDAS = True
except Exception as e:
    HAS_GEOPANDAS = False
    print(f"[Info] geopandas import 실패: {str(e)}")

try:
    import shapefile
    HAS_PYSHP = True
    print(f"[Info] pyshp import 성공 (version: {shapefile.__version__})")
except Exception as e:
    HAS_PYSHP = False
    print(f"[Info] pyshp import 실패: {str(e)}")

from matplotlib.patches import Rectangle, Polygon

# 한글 폰트 설정
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False

# 파일 경로
INPUT_FILE = 'scenario key results.xlsx'
OUTPUT_DIR = 'results/scenario_comparison'

# 지역 코드 매핑
REGION_MAPPING = {
    'SEL': '서울', 'BSN': '부산', 'DGU': '대구', 'ICN': '인천',
    'GWJ': '광주', 'DJN': '대전', 'USN': '울산', 'SJG': '세종',
    'GGD': '경기', 'GWD': '강원', 'CBD': '충북', 'CND': '충남',
    'JBD': '전북', 'JND': '전남', 'GBD': '경북', 'GND': '경남', 'JJD': '제주'
}

# 지역 좌표 (위도, 경도 - 실제 대한민국 좌표 기반)
# 가시성 개선을 위해 일부 지역 위치 조정
REGION_COORDS = {
    'SEL': (127.3, 37.75),   # 서울 (우측상단으로 원 1칸 이동)
    'ICN': (126.7, 37.45),   # 인천
    'GGD': (127.3, 37.35),   # 경기
    'GWD': (128.2, 37.85),   # 강원
    'CBD': (127.7, 36.8),    # 충북
    'CND': (126.8, 36.5),    # 충남
    'SJG': (127.1, 36.33),   # 세종 (좌측하단으로 원 0.5칸 이동)
    'DJN': (127.7, 36.05),   # 대전 (우측하단으로 원 1칸 이동)
    'JBD': (127.1, 35.82),   # 전북
    'JND': (126.7, 34.8),    # 전남
    'GBD': (128.9, 36.5),    # 경북
    'DGU': (128.75, 35.72),  # 대구 (우측하단으로 원 0.5칸 이동)
    'GND': (128.1, 35.25),   # 경남
    'BSN': (129.0, 35.18),   # 부산
    'USN': (129.3, 35.54),   # 울산
    'GWJ': (127.15, 35.16),  # 광주 (우측으로 원 1칸 이동)
    'JJD': (126.53, 33.8)    # 제주 (상단으로 원 1칸 이동)
}

def extract_region_code(bus_name):
    """버스 이름에서 지역 코드 추출"""
    if pd.isna(bus_name):
        return None
    parts = str(bus_name).split('_')
    code = parts[0] if parts else str(bus_name)
    return code if code in REGION_COORDS else None

def plot_transmission_network(ax, scenario_data, year, scenario_name, title):
    """
    송전망 네트워크 그림 생성
    
    Parameters:
    -----------
    ax : matplotlib.axes.Axes
        그림을 그릴 axes
    scenario_data : pd.DataFrame
        시나리오 데이터 (평균조류(MW), 이용률(%) 포함)
    year : int
        연도
    scenario_name : str
        시나리오 이름 (BAU, SC)
    title : str
        그래프 제목
    """
    
    # 1. 배경 그리기 (회색 배경 + 테두리)
    ax.set_facecolor('#e8e8e8')
    
    # 간단한 테두리 그리기
    ax.plot([126.0, 130.0, 130.0, 126.0, 126.0], 
           [33.0, 33.0, 38.5, 38.5, 33.0], 
           'k-', linewidth=1.5, alpha=0.3, zorder=0)
    
    # 2. 지역 노드 그리기 (지역명 포함, 크기 2배)
    for code, (lon, lat) in REGION_COORDS.items():
        # 원 그리기 (크기 2배)
        ax.plot(lon, lat, 'o', markersize=24, color='white', 
               markeredgecolor='darkblue', markeredgewidth=2.5, zorder=5, alpha=0.95)
        # 지역명 표시
        ax.text(lon, lat, REGION_MAPPING.get(code, code), 
               ha='center', va='center', fontsize=9, fontweight='bold', 
               color='darkblue', zorder=6)
    
    # 송전선로 데이터 집계 (같은 노드 간 중복 선로 합산)
    aggregated_lines = {}
    
    for _, row in scenario_data.iterrows():
        bus0_code = extract_region_code(row.get('시작버스'))
        bus1_code = extract_region_code(row.get('종료버스'))
        
        if bus0_code is None or bus1_code is None:
            continue
        if bus0_code not in REGION_COORDS or bus1_code not in REGION_COORDS:
            continue
        
        # 노드 쌍을 정렬하여 방향 무관하게 집계
        node_pair = tuple(sorted([bus0_code, bus1_code]))
        
        flow = row.get('평균조류(MW)', 0)  # 부호 보존
        utilization = row.get('이용률(%)', 0)
        
        # 원래 방향과 정렬된 방향이 다르면 조류 부호 반전
        # 예: 원래 GWD->GGD이지만 정렬 후 (GGD, GWD)가 되면 조류를 음수로
        if (bus0_code, bus1_code) != node_pair:
            flow = -flow  # 방향 반전
        
        if node_pair not in aggregated_lines:
            aggregated_lines[node_pair] = {
                'total_flow': 0,
                'max_utilization': 0,
                'count': 0
            }
        
        aggregated_lines[node_pair]['total_flow'] += flow  # 조정된 조류 합산
        aggregated_lines[node_pair]['max_utilization'] = max(
            aggregated_lines[node_pair]['max_utilization'], 
            utilization
        )
        aggregated_lines[node_pair]['count'] += 1
    
    # 송전선로 그리기
    for (code0, code1), data in aggregated_lines.items():
        lon0, lat0 = REGION_COORDS[code0]
        lon1, lat1 = REGION_COORDS[code1]
        
        total_flow = data['total_flow']
        abs_flow = abs(total_flow)
        max_util = data['max_utilization']
        
        # 선 두께: 조류 크기(절대값)에 비례
        linewidth = max(1.0, min(abs_flow / 150, 10))  # 최소 1, 최대 10
        
        # 선 색상 및 투명도: 이용률에 따라
        if max_util < 40:
            color = 'green'
            alpha = 0.5
        elif max_util < 60:
            color = 'yellowgreen'
            alpha = 0.65
        elif max_util < 80:
            color = 'orange'
            alpha = 0.8
        else:
            color = 'red'
            alpha = 0.95
        
        # 송전선로 그리기
        ax.plot([lon0, lon1], [lat0, lat1], 
               color=color, linewidth=linewidth, alpha=alpha, 
               solid_capstyle='round', zorder=1)
        
        # 화살표와 송전량 표시 (큰 조류만)
        if abs_flow > 50:  # 50MW 이상만 표시
            # 조류의 부호에 따라 방향 결정
            # 양수: code0 -> code1, 음수: code1 -> code0
            if total_flow > 0:
                # 정방향: code0 -> code1
                from_lon, from_lat = lon0, lat0
                to_lon, to_lat = lon1, lat1
            else:
                # 역방향: code1 -> code0
                from_lon, from_lat = lon1, lat1
                to_lon, to_lat = lon0, lat0
            
            # 선의 방향 벡터
            dx = to_lon - from_lon
            dy = to_lat - from_lat
            length = np.sqrt(dx**2 + dy**2)
            
            if length > 0:
                # 정규화된 방향 벡터
                dx_norm = dx / length
                dy_norm = dy / length
                
                # 화살표 길이 설정
                arrow_length = 0.3
                
                # 화살표 시작점: 목적지 버스에서 역방향으로 떨어진 위치
                arrow_start_lon = to_lon - dx_norm * (arrow_length + 0.2)
                arrow_start_lat = to_lat - dy_norm * (arrow_length + 0.2)
                
                # 화살표 벡터 (시작점에서 목적지 버스 방향으로)
                arrow_dx = dx_norm * arrow_length
                arrow_dy = dy_norm * arrow_length
                
                # 화살표 그리기 (머리가 목적지 버스에 닿도록, 검은색)
                ax.arrow(arrow_start_lon, arrow_start_lat, arrow_dx, arrow_dy, 
                        head_width=0.06, head_length=0.05, 
                        fc='black', ec='black', alpha=0.8, 
                        linewidth=1.2, zorder=4)
                
                # 송전량 텍스트 표시 (화살표 옆) - 절대값으로 표시
                # 화살표 중간지점 계산
                arrow_mid_lon = arrow_start_lon + arrow_dx/2
                arrow_mid_lat = arrow_start_lat + arrow_dy/2
                
                # 텍스트를 화살표 옆(수직 방향)에 표시
                perp_dx = -dy_norm * 0.12
                perp_dy = dx_norm * 0.12
                text_lon = arrow_mid_lon + perp_dx
                text_lat = arrow_mid_lat + perp_dy
                
                ax.text(text_lon, text_lat, f'{abs_flow:.0f}MW', 
                       ha='center', va='center', fontsize=7, 
                       bbox=dict(boxstyle='round,pad=0.3', facecolor='white', 
                                edgecolor='gray', alpha=0.8),
                       zorder=4)
    
    # 범례 (우측 하단)
    legend_elements = [
        mpatches.Patch(facecolor='green', alpha=0.5, label='이용률 < 40%'),
        mpatches.Patch(facecolor='yellowgreen', alpha=0.65, label='이용률 40-60%'),
        mpatches.Patch(facecolor='orange', alpha=0.8, label='이용률 60-80%'),
        mpatches.Patch(facecolor='red', alpha=0.95, label='이용률 ≥ 80%'),
        plt.Line2D([0], [0], color='gray', linewidth=2, label='선 두께 = 조류 크기')
    ]
    ax.legend(handles=legend_elements, loc='lower right', fontsize=10, 
             framealpha=0.95, edgecolor='black', fancybox=True, shadow=True)
    
    # 축 설정
    ax.set_xlim(126.0, 130.0)
    ax.set_ylim(33.0, 38.5)
    ax.set_aspect('equal')
    
    # 축 라벨과 눈금 제거하되, 배경은 유지
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.set_xticks([])
    ax.set_yticks([])
    ax.grid(False)
    
    # 제목 (axis가 꺼져있으므로 text로 표시)
    ax.text(0.5, 1.02, title, transform=ax.transAxes, 
           fontsize=15, fontweight='bold', ha='center')
    
    # 통계 정보 표시 (좌측 상단)
    stats_text = f"송전선로: {len(aggregated_lines)}개\n"
    stats_text += f"평균 조류: {sum(d['total_flow'] for d in aggregated_lines.values())/len(aggregated_lines):.0f} MW"
    ax.text(0.02, 0.98, stats_text, transform=ax.transAxes, 
           fontsize=10, verticalalignment='top', 
           bbox=dict(boxstyle='round', facecolor='lightyellow', alpha=0.9, 
                    edgecolor='black', linewidth=1.5))

def create_scenario_comparison():
    """시나리오별 송전망 비교 시각화 생성"""
    
    print(f"\n{'='*70}")
    print(f"시나리오별 송전망 비교 시각화")
    print(f"{'='*70}")
    
    # 입력 파일 확인
    if not os.path.exists(INPUT_FILE):
        print(f"\n[오류] '{INPUT_FILE}' 파일을 찾을 수 없습니다.")
        print(f"현재 디렉토리에 '{INPUT_FILE}' 파일이 있는지 확인하세요.")
        return False
    
    print(f"입력 파일: {INPUT_FILE}")
    
    # 출력 디렉토리 생성
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"출력 디렉토리: {OUTPUT_DIR}/")
    
    # 엑셀 파일 읽기 (멀티레벨 헤더)
    print(f"\n데이터 로딩 중...")
    try:
        # 첫 3행을 헤더로 사용 (시나리오, 연도, 항목)
        df_raw = pd.read_excel(INPUT_FILE, sheet_name='네트워크 혼잡', header=[0, 1, 2])
        
        # 컬럼명을 단일 레벨로 병합 (예: "BAU_2030_평균조류(MW)")
        new_columns = []
        for col in df_raw.columns:
            # 튜플 형태의 멀티인덱스를 문자열로 결합
            col_parts = [str(c).strip() for c in col if not str(c).startswith('Unnamed')]
            if col_parts:
                new_columns.append('_'.join(col_parts))
            else:
                # Unnamed만 있는 경우는 시작버스, 종료버스 등
                new_columns.append(str(col[0]))
        
        df_raw.columns = new_columns
        
        # 시작버스, 종료버스 컬럼 찾기
        bus_cols = [col for col in df_raw.columns if '버스' in str(col) or 'bus' in str(col).lower()]
        if len(bus_cols) < 2:
            # 처음 두 컬럼을 시작버스, 종료버스로 가정
            df = df_raw.copy()
            df.columns = ['시작버스', '종료버스'] + list(df.columns[2:])
        else:
            df = df_raw.copy()
            # 첫 번째 버스 컬럼을 시작버스로, 두 번째를 종료버스로
            col_rename = {bus_cols[0]: '시작버스', bus_cols[1]: '종료버스'}
            df.rename(columns=col_rename, inplace=True)
        
        print(f"[OK] 데이터 로드 완료: {len(df)}개 송전선로")
        
        # 컬럼명 출력 (디버깅용)
        print(f"\n컬럼 목록 (병합 후):")
        for i, col in enumerate(df.columns, 1):
            print(f"  {i}. {col}")
            
    except Exception as e:
        print(f"[오류] 데이터 로드 실패 - {str(e)}")
        import traceback
        traceback.print_exc()
        return False
    
    # 연도별 처리
    years = [2030, 2038]
    scenarios = ['BAU', 'SC']
    
    for year in years:
        print(f"\n{'='*70}")
        print(f"{year}년 시나리오 비교 처리 중...")
        print(f"{'='*70}")
        
        # 각 시나리오별 데이터 준비
        scenario_datasets = {}
        
        for scenario in scenarios:
            # 해당 연도, 시나리오에 맞는 컬럼 찾기
            matching_cols = [col for col in df.columns 
                           if str(year) in str(col) and scenario in str(col)]
            
            if not matching_cols:
                print(f"[경고] {year}년 {scenario} 데이터를 찾을 수 없습니다.")
                continue
            
            # 데이터 복사
            scenario_data = df[['시작버스', '종료버스']].copy()
            
            # 평균조류와 이용률 컬럼 찾기
            flow_col = None
            util_col = None
            
            for col in matching_cols:
                col_str = str(col).lower()
                if '조류' in col_str or 'flow' in col_str:
                    flow_col = col
                elif '이용률' in col_str or 'utilization' in col_str:
                    util_col = col
            
            if flow_col:
                scenario_data['평균조류(MW)'] = df[flow_col].abs()
                print(f"  [OK] {scenario}: 조류 데이터 ({flow_col})")
            else:
                print(f"  [경고] {scenario}: 조류 데이터를 찾을 수 없습니다.")
                continue
            
            if util_col:
                scenario_data['이용률(%)'] = df[util_col]
                print(f"  [OK] {scenario}: 이용률 데이터 ({util_col})")
            else:
                print(f"  [경고] {scenario}: 이용률 데이터를 찾을 수 없습니다.")
                scenario_data['이용률(%)'] = 0
            
            scenario_datasets[scenario] = scenario_data
        
        # 두 시나리오가 모두 있는 경우에만 그림 생성
        if 'BAU' not in scenario_datasets or 'SC' not in scenario_datasets:
            print(f"[경고] {year}년: BAU 또는 SC 데이터가 부족하여 건너뜁니다.")
            continue
        
        # 그림 생성 (1x2 레이아웃, 10% 확대)
        print(f"\n{year}년 송전망 비교 그림 생성 중...")
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(26.4, 12.1))
        
        plot_transmission_network(
            ax1, scenario_datasets['BAU'], year, 'BAU',
            f'{year}년 송전망 현황 - BAU 시나리오'
        )
        
        plot_transmission_network(
            ax2, scenario_datasets['SC'], year, 'SC',
            f'{year}년 송전망 현황 - SC 시나리오 (유연성 강화)'
        )
        
        # 전체 제목
        fig.suptitle(
            f'{year}년 시나리오별 송전망 비교\n' + 
            '(선 두께: 조류 크기 / 선 색상: 이용률)',
            fontsize=18, fontweight='bold', y=0.98
        )
        
        plt.tight_layout(rect=[0, 0, 1, 0.96])
        
        # 저장
        output_file = f'{OUTPUT_DIR}/transmission_comparison_{year}.png'
        plt.savefig(output_file, dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()
        
        print(f"[OK] 저장 완료: {output_file}")
    
    print(f"\n{'='*70}")
    print(f"[OK] 모든 시각화 완료!")
    print(f"결과 파일:")
    for year in years:
        output_file = f'{OUTPUT_DIR}/transmission_comparison_{year}.png'
        if os.path.exists(output_file):
            print(f"  - {output_file}")
    print(f"{'='*70}\n")
    
    return True

def create_hourly_animation():
    """
    시간대별 송전망 이용 현황 애니메이션 생성
    
    network1: 2038 BAU 시나리오의 시간별 line 이용 데이터
    network2: 2038 SC 시나리오의 시간별 line 이용 데이터
    
    각 시간대(1~24시)별로 통계를 내어 24개 프레임의 애니메이션 생성
    """
    
    print(f"\n{'='*70}")
    print(f"시간대별 송전망 애니메이션 생성")
    print(f"{'='*70}")
    
    # 입력 파일 확인
    if not os.path.exists(INPUT_FILE):
        print(f"\n[오류] '{INPUT_FILE}' 파일을 찾을 수 없습니다.")
        return False
    
    print(f"입력 파일: {INPUT_FILE}")
    
    # 출력 디렉토리 생성
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"출력 디렉토리: {OUTPUT_DIR}/")
    
    # network1, network2 데이터 읽기
    print(f"\n데이터 로딩 중...")
    try:
        # network1: 2038 BAU (1행=선로명, 2행부터=시간별 데이터)
        df_network1_raw = pd.read_excel(INPUT_FILE, sheet_name='network1')
        print(f"[OK] network1 (2038 BAU) 원본 로드 완료: {df_network1_raw.shape}")
        
        # network2: 2038 SC
        df_network2_raw = pd.read_excel(INPUT_FILE, sheet_name='network2')
        print(f"[OK] network2 (2038 SC) 원본 로드 완료: {df_network2_raw.shape}")
        
        # 네트워크 혼잡 시트에서 선로명-버스 매핑 정보 가져오기
        try:
            # 헤더 없이 직접 읽기 (4행부터 데이터, 1-3행은 헤더)
            df_congestion = pd.read_excel(INPUT_FILE, sheet_name='네트워크 혼잡', skiprows=3)
            
            # 첫 3개 컬럼: 선로명(A), 시작버스(B), 종료버스(C)
            line_bus_mapping = {}
            for _, row in df_congestion.iterrows():
                line_name = row.iloc[0]  # A열: 선로명
                bus0 = row.iloc[1]       # B열: 시작버스
                bus1 = row.iloc[2]       # C열: 종료버스
                
                if not pd.isna(line_name):
                    line_bus_mapping[str(line_name).strip()] = {
                        'bus0': str(bus0).strip() if not pd.isna(bus0) else None,
                        'bus1': str(bus1).strip() if not pd.isna(bus1) else None
                    }
            
            print(f"[OK] 네트워크 혼잡 시트 로드 완료: {len(line_bus_mapping)}개 선로 매핑")
            
            # 샘플 출력
            if len(line_bus_mapping) > 0:
                sample_key = list(line_bus_mapping.keys())[0]
                sample_val = line_bus_mapping[sample_key]
                print(f"    [샘플] {sample_key}: {sample_val['bus0']} -> {sample_val['bus1']}")
                
        except Exception as e:
            print(f"[오류] 네트워크 혼잡 시트 로드 실패: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
        
        # regional_el1, regional_el2 시트 로드 (지역별 수요 데이터)
        try:
            df_regional_el1 = pd.read_excel(INPUT_FILE, sheet_name='regional_el1')
            print(f"[OK] regional_el1 (2038 BAU 수요) 로드 완료: {df_regional_el1.shape}")
            
            df_regional_el2 = pd.read_excel(INPUT_FILE, sheet_name='regional_el2')
            print(f"[OK] regional_el2 (2038 SC 수요) 로드 완료: {df_regional_el2.shape}")
        except Exception as e:
            print(f"[경고] 지역별 수요 데이터 로드 실패: {str(e)}")
            df_regional_el1 = None
            df_regional_el2 = None
        
    except Exception as e:
        print(f"[오류] 데이터 로딩 실패: {str(e)}")
        import traceback
        traceback.print_exc()
        return False
    
    # 시간대별 통계 계산
    print(f"\n시간대별 통계 계산 중...")
    hourly_stats = {}
    
    for scenario_name, df_raw in [('BAU', df_network1_raw), ('SC', df_network2_raw)]:
        print(f"  [처리 중] 2038 {scenario_name}")
        
        # 데이터 구조 확인
        print(f"    원본 데이터 크기: {df_raw.shape}")
        print(f"    전체 컬럼 수: {len(df_raw.columns)}")
        print(f"    첫 5개 컬럼(선로명): {df_raw.columns[:5].tolist()}")
        
        # 각 시간대(0~23)별 평균 계산
        hourly_data = {hour: [] for hour in range(24)}
        
        processed_lines = 0
        skipped_lines = 0
        skipped_not_in_mapping = 0
        
        # 각 컬럼(선로)별로 처리
        for col_idx, line_name in enumerate(df_raw.columns):
            line_name_str = str(line_name).strip()
            
            # 네트워크 혼잡 시트의 매핑에서 버스 정보 가져오기
            if line_name_str not in line_bus_mapping:
                skipped_not_in_mapping += 1
                if col_idx < 3:
                    print(f"    [경고] 선로 {col_idx} ({line_name_str}): 매핑에 없음")
                continue
            
            bus_info = line_bus_mapping[line_name_str]
            bus0 = bus_info['bus0']
            bus1 = bus_info['bus1']
            
            # 첫 번째 선로만 디버깅 출력
            if col_idx == 0:
                print(f"    [디버깅] 첫 번째 송전선:")
                print(f"      선로명: {line_name_str}")
                print(f"      매핑 결과 -> 시작버스: {bus0}, 종료버스: {bus1}")
                print(f"      시작버스 지역코드: {extract_region_code(bus0)}")
                print(f"      종료버스 지역코드: {extract_region_code(bus1)}")
            
            if bus0 is None or bus1 is None:
                skipped_lines += 1
                if col_idx < 3:
                    print(f"    [경고] 선로 {col_idx} ({line_name_str}): 버스 정보 없음")
                continue
            
            # 해당 컬럼의 시간별 데이터 (8760개)
            time_values = df_raw[line_name].values
            
            if len(time_values) >= 8760:
                # 8760개 또는 그 이상의 데이터만 처리
                for hour in range(24):
                    # 해당 시간대의 모든 날 데이터 추출 (hour, hour+24, hour+48, ...)
                    hour_indices = list(range(hour, min(8760, len(time_values)), 24))
                    hour_values = [time_values[i] for i in hour_indices if not pd.isna(time_values[i]) and not np.isnan(time_values[i])]
                    
                    if hour_values:
                        hour_avg = np.mean(hour_values)
                        
                        hourly_data[hour].append({
                            'line': line_name_str,
                            'bus0': bus0,
                            'bus1': bus1,
                            'flow': hour_avg  # 부호 보존!
                        })
                
                processed_lines += 1
            else:
                skipped_lines += 1
                if col_idx < 3:
                    print(f"    [경고] 선로 {col_idx} ({line_name}): 시간 데이터 부족 ({len(time_values)}개)")
        
        hourly_stats[scenario_name] = hourly_data
        print(f"    [OK] 24시간 통계 계산 완료")
        print(f"    처리된 송전선: {processed_lines}개")
        print(f"    건너뛴 송전선: {skipped_lines}개")
        print(f"    매핑에 없는 선로: {skipped_not_in_mapping}개")
        
        # 0시 데이터 샘플 확인
        if 0 in hourly_data and len(hourly_data[0]) > 0:
            print(f"    [샘플] 0시 데이터 개수: {len(hourly_data[0])}개")
            sample = hourly_data[0][0]
            print(f"      첫 번째 송전선: {sample['bus0']} -> {sample['bus1']}, 평균조류: {sample['flow']:.2f}MW")
    
    # 지역별 수요 데이터 시간별 집계
    print(f"\n지역별 수요 데이터 시간별 집계 중...")
    hourly_demand = {}
    
    for scenario_name, df_demand in [('BAU', df_regional_el1), ('SC', df_regional_el2)]:
        if df_demand is None:
            print(f"  [건너뛰기] {scenario_name} 수요 데이터 없음")
            hourly_demand[scenario_name] = {hour: {} for hour in range(24)}
            continue
        
        print(f"  [처리 중] {scenario_name} 수요 데이터")
        print(f"    컬럼 목록: {list(df_demand.columns[:5])}...")  # 처음 5개 컬럼만
        hourly_demand_data = {hour: {} for hour in range(24)}
        
        # 데이터프레임의 컬럼이 지역명, 행이 시간별 데이터라고 가정
        for region_col in df_demand.columns:
            region_name_str = str(region_col).strip()
            region_code = extract_region_code(region_name_str)
            
            # SJG(세종) 특별 처리: 컬럼명에 '세종'이 포함되어 있으면 SJG로 매핑
            if region_code is None and ('세종' in region_name_str or 'SJN' in region_name_str):
                region_code = 'SJG'
            
            if region_code is None:
                continue
            
            # 해당 지역의 시간별 데이터 (8760개)
            time_values = df_demand[region_col].values
            
            if len(time_values) >= 8760:
                for hour in range(24):
                    # 해당 시간대의 모든 날 데이터 추출
                    hour_indices = list(range(hour, min(8760, len(time_values)), 24))
                    hour_values = [time_values[i] for i in hour_indices if not pd.isna(time_values[i])]
                    
                    if hour_values:
                        hour_avg_demand = np.mean(hour_values)
                        hourly_demand_data[hour][region_code] = hour_avg_demand
        
        hourly_demand[scenario_name] = hourly_demand_data
        print(f"    [OK] {scenario_name} 수요 데이터 집계 완료 (지역 수: {len(hourly_demand_data[0])}개)")
        
        # 샘플 출력
        if len(hourly_demand_data[0]) > 0:
            sample_region = list(hourly_demand_data[0].keys())[0]
            print(f"    [샘플] 0시 {sample_region} 수요: {hourly_demand_data[0][sample_region]:.2f}MW")
    
    # 애니메이션 생성 (BAU와 SC를 나란히 배치)
    print(f"\n애니메이션 생성 중...")
    print(f"  [생성 중] 2038 BAU vs SC 비교 애니메이션")
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(26.4, 12.1))
    
    def draw_hour_network(ax, scenario_name, hour):
        """특정 시간대의 송전망 그리기"""
        
        # 배경 그리기
        ax.set_facecolor('#e8e8e8')
        ax.plot([126.0, 130.0, 130.0, 126.0, 126.0], 
               [33.0, 33.0, 38.5, 38.5, 33.0], 
               'k-', linewidth=1.5, alpha=0.3, zorder=0)
        
        # 해당 시간대 데이터로 송전선 그리기
        hour_data = hourly_stats[scenario_name][hour]
        
        # 먼저 송전선을 집계하여 각 지역의 송전/수전량 계산
        region_flows = {}  # 각 지역의 순수입(+)/순수출(-) 계산용
        for code in REGION_COORDS.keys():
            region_flows[code] = {'import': 0, 'export': 0}  # 수입(수전), 수출(송전)
        
        # 각 프레임별 데이터 개수 확인 (처음 3개만 출력)
        if hour < 3 or hour == 23:
            print(f"    [디버깅] {scenario_name} {hour}시 - 데이터: {len(hour_data)}개 송전선")
            if len(hour_data) > 0:
                sample_flow = hour_data[0]['flow']
                print(f"      첫 번째 선로 조류: {sample_flow:.2f}MW")
        
        # 노드 쌍별로 집계
        aggregated_lines = {}
        skipped_no_code = 0
        skipped_not_in_coords = 0
        
        for item in hour_data:
            bus0_code = extract_region_code(item['bus0'])
            bus1_code = extract_region_code(item['bus1'])
            
            if bus0_code is None or bus1_code is None:
                skipped_no_code += 1
                continue
            if bus0_code not in REGION_COORDS or bus1_code not in REGION_COORDS:
                skipped_not_in_coords += 1
                continue
            
            node_pair = tuple(sorted([bus0_code, bus1_code]))
            flow = item['flow']
            
            # 원래 방향과 정렬된 방향이 다르면 조류 부호 반전
            if (bus0_code, bus1_code) != node_pair:
                flow = -flow  # 방향 반전
            
            if node_pair not in aggregated_lines:
                aggregated_lines[node_pair] = {'total_flow': 0, 'count': 0}
            
            aggregated_lines[node_pair]['total_flow'] += flow
            aggregated_lines[node_pair]['count'] += 1
        
        # 각 프레임 집계 결과 출력 (처음 3개만 출력)
        if hour < 3 or hour == 23:
            print(f"    [디버깅] {scenario_name} {hour}시 집계 결과: {len(aggregated_lines)}개 송전선")
            if len(aggregated_lines) > 0:
                sample_pair = list(aggregated_lines.keys())[0]
                print(f"      첫 번째 송전선: {sample_pair}, 조류: {aggregated_lines[sample_pair]['total_flow']:.2f}MW")
            
            # 사용자가 언급한 특정 선로들 확인 (0시만)
            if hour == 0:
                check_pairs = [
                    ('GBD', 'GWD'), ('GGD', 'ICN'), ('GBD', 'GND'), ('BSN', 'GND'), 
                    ('GND', 'USN'), ('GND', 'JND'), ('GWJ', 'JND'), ('DJN', 'SJG'),
                    ('CND', 'DJN'), ('CBD', 'DJN'), ('BSN', 'USN'), ('CBD', 'CND')
                ]
                print(f"    [Check specific lines at 0:00]")
                for pair in check_pairs:
                    if pair in aggregated_lines:
                        flow = aggregated_lines[pair]['total_flow']
                        if flow > 0:
                            print(f"      {pair[0]} -> {pair[1]}: +{flow:.1f}MW")
                        else:
                            print(f"      {pair[1]} -> {pair[0]}: {abs(flow):.1f}MW (reversed)")
        
        # 송전선로 그리기 (먼저 송전/수전량 계산)
        max_flow = max([abs(d['total_flow']) for d in aggregated_lines.values()], default=1)
        
        for (code0, code1), data in aggregated_lines.items():
            lon0, lat0 = REGION_COORDS[code0]
            lon1, lat1 = REGION_COORDS[code1]
            
            total_flow = data['total_flow']
            abs_flow = abs(total_flow)
            
            # 각 지역의 송전/수전량 누적 (양수: code0->code1 방향)
            if total_flow > 0:
                # code0에서 code1로 흐름 (code0은 수출, code1은 수입)
                region_flows[code0]['export'] += abs_flow
                region_flows[code1]['import'] += abs_flow
            else:
                # code1에서 code0로 흐름 (code1은 수출, code0은 수입)
                region_flows[code1]['export'] += abs_flow
                region_flows[code0]['import'] += abs_flow
            
            # 선 두께: 조류 크기(절대값)에 비례
            linewidth = max(1.0, min(abs_flow / 150, 10))
            
            # 선 색상: 유량 크기(절대값)에 따라
            if abs_flow < max_flow * 0.3:
                color = 'green'
                alpha = 0.5
            elif abs_flow < max_flow * 0.6:
                color = 'yellowgreen'
                alpha = 0.65
            else:
                color = 'orange'
                alpha = 0.8
            
            # 송전선로 그리기
            ax.plot([lon0, lon1], [lat0, lat1], 
                   color=color, linewidth=linewidth, alpha=alpha, 
                   solid_capstyle='round', zorder=1)
            
            # 화살표 (큰 조류만)
            if abs_flow > 50:
                # 조류의 부호에 따라 방향 결정
                # 양수: code0 -> code1, 음수: code1 -> code0
                if total_flow > 0:
                    # 정방향: code0 -> code1
                    from_lon, from_lat = lon0, lat0
                    to_lon, to_lat = lon1, lat1
                    from_code, to_code = code0, code1
                else:
                    # 역방향: code1 -> code0
                    from_lon, from_lat = lon1, lat1
                    to_lon, to_lat = lon0, lat0
                    from_code, to_code = code1, code0
                
                dx = to_lon - from_lon
                dy = to_lat - from_lat
                length = np.sqrt(dx**2 + dy**2)
                
                if length > 0:
                    dx_norm = dx / length
                    dy_norm = dy / length
                    
                    arrow_length = 0.3
                    # 화살표 머리가 목적지 버스에 닿도록
                    arrow_start_lon = to_lon - dx_norm * (arrow_length + 0.2)
                    arrow_start_lat = to_lat - dy_norm * (arrow_length + 0.2)
                    
                    arrow_dx = dx_norm * arrow_length
                    arrow_dy = dy_norm * arrow_length
                    
                    ax.arrow(arrow_start_lon, arrow_start_lat, arrow_dx, arrow_dy, 
                            head_width=0.06, head_length=0.05, 
                            fc='black', ec='black', alpha=0.8, 
                            linewidth=1.2, zorder=4)
                    
                    # 화살표 옆에 조류 수치 표시
                    arrow_mid_lon = arrow_start_lon + arrow_dx/2
                    arrow_mid_lat = arrow_start_lat + arrow_dy/2
                    
                    # 텍스트를 화살표 옆(수직 방향)에 표시
                    perp_dx = -dy_norm * 0.1
                    perp_dy = dx_norm * 0.1
                    text_lon = arrow_mid_lon + perp_dx
                    text_lat = arrow_mid_lat + perp_dy
                    
                    ax.text(text_lon, text_lat, f'{abs_flow:.0f}', 
                           ha='center', va='center', fontsize=6, 
                           bbox=dict(boxstyle='round,pad=0.2', facecolor='white', 
                                    edgecolor='gray', alpha=0.8),
                           zorder=5)
        
        # 지역 노드 그리기 (송전선 위에 표시)
        for code, (lon, lat) in REGION_COORDS.items():
            ax.plot(lon, lat, 'o', markersize=24, color='white', 
                   markeredgecolor='darkblue', markeredgewidth=2.5, zorder=5, alpha=0.95)
            
            # 지역명
            ax.text(lon, lat, REGION_MAPPING.get(code, code), 
                   ha='center', va='center', fontsize=9, fontweight='bold', 
                   color='darkblue', zorder=6)
            
            # 수요, 송전, 수전 정보 표시 (노드 우측)
            demand = hourly_demand[scenario_name][hour].get(code, 0)
            import_flow = region_flows[code]['import']
            export_flow = region_flows[code]['export']
            
            info_text = f"수요:{demand:.0f}\n수전:{import_flow:.0f}\n송전:{export_flow:.0f}"
            ax.text(lon, lat -0.15, info_text, 
                   ha='center', va='top', fontsize=5, 
                   bbox=dict(boxstyle='round,pad=0.3', facecolor='lightyellow', 
                            edgecolor='gray', alpha=0.95),
                   zorder=6)
        
        # 축 설정
        ax.set_xlim(126.0, 130.0)
        ax.set_ylim(33.0, 38.5)
        ax.set_aspect('equal')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['bottom'].set_visible(False)
        ax.spines['left'].set_visible(False)
        ax.set_xticks([])
        ax.set_yticks([])
        ax.grid(False)
        
        # 서브플롯 제목
        title = f'2038 {scenario_name} 시나리오'
        ax.text(0.5, 1.02, title, transform=ax.transAxes, 
               fontsize=15, fontweight='bold', ha='center')
        
        # 시각 정보 (중앙 상단)
        time_text = f"{hour:02d}:00"
        ax.text(0.5, 0.96, time_text, transform=ax.transAxes, 
               fontsize=14, fontweight='bold', ha='center', va='top',
               bbox=dict(boxstyle='round,pad=0.5', facecolor='lightyellow', alpha=0.95, 
                        edgecolor='darkblue', linewidth=2))
    
    # 애니메이션 업데이트 함수
    def update(hour):
        print(f"  [프레임] {hour:02d}시 그리는 중...")
        
        # 두 서브플롯 모두 초기화
        ax1.clear()
        ax2.clear()
        
        # BAU 그리기 (왼쪽)
        draw_hour_network(ax1, 'BAU', hour)
        
        # SC 그리기 (오른쪽)
        draw_hour_network(ax2, 'SC', hour)
        
        # 전체 제목
        fig.suptitle(
            f'2038년 시나리오별 송전망 흐름 비교 - {hour:02d}시\n' +
            '(선 두께: 조류 크기 / 선 색상: 유량 크기)',
            fontsize=18, fontweight='bold', y=0.98
        )
    
    # 애니메이션 생성
    print(f"  [애니메이션] FuncAnimation 객체 생성 중...")
    anim = FuncAnimation(fig, update, frames=24, interval=1000, repeat=True, blit=False)
    
    # GIF로 저장
    output_file = f'{OUTPUT_DIR}/transmission_hourly_2038_comparison.gif'
    print(f"  [저장 중] {output_file}")
    print(f"  총 24프레임 저장 중... (시간이 걸릴 수 있습니다)")
    
    plt.tight_layout(rect=[0, 0, 1, 0.96])
    
    try:
        writer = PillowWriter(fps=1)  # 1초당 1프레임
        anim.save(output_file, writer=writer, dpi=100, savefig_kwargs={'facecolor': 'white'})
        print(f"  [OK] 저장 완료: {output_file}")
        
        # 파일 크기 확인
        file_size = os.path.getsize(output_file) / (1024 * 1024)  # MB
        print(f"  파일 크기: {file_size:.1f} MB")
        
    except Exception as e:
        print(f"  [오류] 저장 실패: {str(e)}")
        import traceback
        traceback.print_exc()
    
    plt.close()
    
    print(f"\n{'='*70}")
    print(f"[OK] 애니메이션 생성 완료!")
    print(f"결과 파일:")
    output_file = f'{OUTPUT_DIR}/transmission_hourly_2038_comparison.gif'
    if os.path.exists(output_file):
        print(f"  - {output_file}")
    print(f"{'='*70}\n")
    
    return True

if __name__ == '__main__':
    import sys
    
    # 기존 시나리오 비교 실행
    success1 = create_scenario_comparison()
    
    # 시간대별 애니메이션 실행
    success2 = create_hourly_animation()
    
    sys.exit(0 if (success1 and success2) else 1)

