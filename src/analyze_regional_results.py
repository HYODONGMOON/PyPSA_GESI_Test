import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.colors import LinearSegmentedColormap
import os
import glob
import folium
from folium.plugins import MarkerCluster
from branca.colormap import linear
from branca.element import Figure
import networkx as nx
import warnings
import pypsa
import re
from matplotlib import font_manager
import sys
import traceback

warnings.filterwarnings('ignore')

# 폰트 설정
try:
    font_path = 'KoPubWorld Dotum Medium.ttf'
    if os.path.exists(font_path):
        font_manager.fontManager.addfont(font_path)
        plt.rcParams['font.family'] = 'KoPubWorld Dotum Medium'
    else:
        print(f"경고: 폰트 파일을 찾을 수 없습니다: {font_path}")
        # 시스템 기본 글꼴 사용
        plt.rcParams['font.family'] = 'Malgun Gothic' if sys.platform == 'win32' else 'AppleGothic' if sys.platform == 'darwin' else 'NanumGothic'
except Exception as e:
    print(f"폰트 설정 중 오류 발생: {str(e)}")
    # 기본 폰트 사용
    plt.rcParams['font.family'] = 'sans-serif'

# 결과 파일이 저장된 디렉토리
RESULTS_DIR = 'results'

def get_latest_timestamp():
    """최신 타임스탬프 가져오기"""
    timestamp_file = os.path.join(RESULTS_DIR, "latest_timestamp.txt")
    if os.path.exists(timestamp_file):
        with open(timestamp_file, 'r') as f:
            timestamp = f.read().strip()
            return timestamp
    
    # 파일이 없는 경우 최신 결과 폴더 검색
    result_dirs = glob.glob(os.path.join(RESULTS_DIR, "[0-9]" * 8 + "_" + "[0-9]" * 6))
    if result_dirs:
        # 폴더 수정 시간으로 정렬
        latest_dir = max(result_dirs, key=os.path.getmtime)
        return os.path.basename(latest_dir)
    
    # 아무것도 없는 경우 예외 발생
    raise FileNotFoundError("최신 타임스탬프 정보를 찾을 수 없습니다.")

def get_results_dir():
    """최신 결과 디렉토리 경로 가져오기"""
    try:
        timestamp = get_latest_timestamp()
        results_dir = os.path.join(RESULTS_DIR, timestamp)
        if os.path.exists(results_dir):
            return results_dir
    except Exception as e:
        print(f"최신 결과 디렉토리를 찾는 중 오류 발생: {str(e)}")
    
    # 실패한 경우 기본 폴더 반환
    return RESULTS_DIR

def get_latest_result_file(pattern='optimization_result_*.nc'):
    """최신 최적화 결과 파일 가져오기"""
    # 먼저 최신 결과 폴더에서 검색
    results_dir = get_results_dir()
    result_files = glob.glob(os.path.join(results_dir, pattern))
    
    # 없으면 전체 results 폴더에서 검색
    if not result_files:
        result_files = glob.glob(os.path.join(RESULTS_DIR, pattern))
        if not result_files:
            result_files = glob.glob(os.path.join(RESULTS_DIR, "*", pattern))
    
    if not result_files:
        raise FileNotFoundError(f"결과 파일이 없습니다: {pattern}")
    
    # 파일 수정 시간으로 정렬
    latest_file = max(result_files, key=os.path.getmtime)
    return latest_file

def get_latest_generator_output():
    """최신 발전기 출력 결과 파일 가져오기"""
    # 먼저 최신 결과 폴더에서 검색
    results_dir = get_results_dir()
    result_files = glob.glob(os.path.join(results_dir, 'optimization_result_*_generator_output.csv'))
    
    # 없으면 전체 results 폴더에서 검색
    if not result_files:
        result_files = glob.glob(os.path.join(RESULTS_DIR, 'optimization_result_*_generator_output.csv'))
        if not result_files:
            result_files = glob.glob(os.path.join(RESULTS_DIR, "*", 'optimization_result_*_generator_output.csv'))
    
    if not result_files:
        raise FileNotFoundError("발전기 출력 결과 파일이 없습니다.")
    
    # 파일 수정 시간으로 정렬
    latest_file = max(result_files, key=os.path.getmtime)
    return latest_file

def get_latest_line_usage():
    """최신 라인 사용률 결과 파일 가져오기"""
    # 먼저 최신 결과 폴더에서 검색
    results_dir = get_results_dir()
    result_files = glob.glob(os.path.join(results_dir, 'optimization_result_*_line_usage.csv'))
    
    # 없으면 전체 results 폴더에서 검색
    if not result_files:
        result_files = glob.glob(os.path.join(RESULTS_DIR, 'optimization_result_*_line_usage.csv'))
        if not result_files:
            result_files = glob.glob(os.path.join(RESULTS_DIR, "*", 'optimization_result_*_line_usage.csv'))
    
    if not result_files:
        raise FileNotFoundError("라인 사용률 결과 파일이 없습니다.")
    
    # 파일 수정 시간으로 정렬
    latest_file = max(result_files, key=os.path.getmtime)
    return latest_file

def get_latest_load_data():
    """최신 부하 데이터 파일 가져오기"""
    # 먼저 최신 결과 폴더에서 검색
    results_dir = get_results_dir()
    result_files = glob.glob(os.path.join(results_dir, 'optimization_result_*_load.csv'))
    
    # 없으면 전체 results 폴더에서 검색
    if not result_files:
        result_files = glob.glob(os.path.join(RESULTS_DIR, 'optimization_result_*_load.csv'))
        if not result_files:
            result_files = glob.glob(os.path.join(RESULTS_DIR, "*", 'optimization_result_*_load.csv'))
    
    if not result_files:
        raise FileNotFoundError("부하 데이터 파일이 없습니다.")
    
    # 파일 수정 시간으로 정렬
    latest_file = max(result_files, key=os.path.getmtime)
    return latest_file

def extract_region_code(name):
    """이름에서 지역 코드 추출하기 (예: BSN_H -> BSN)"""
    if '_' in name:
        return name.split('_')[0]
    return name

def get_korean_region_name(code):
    """지역 코드에 대한 한국어 이름 반환"""
    region_names = {
        'BSN': '부산광역시',
        'CBD': '충청북도',
        'CND': '충청남도',
        'DGU': '대구광역시',
        'DJN': '대전광역시',
        'GBD': '경상북도',
        'GGD': '경기도',
        'GND': '경상남도',
        'GWD': '강원도',
        'GWJ': '광주광역시',
        'ICN': '인천광역시',
        'JBD': '전라북도',
        'JJD': '제주특별자치도',
        'JND': '전라남도',
        'SEL': '서울특별시',
        'SJN': '세종특별자치시',
        'USN': '울산광역시'
    }
    return region_names.get(code, code)

def analyze_regional_energy_balance(output_dir=None):
    """지역별 에너지 균형을 분석하고 결과를 CSV 파일로 저장합니다."""
    print("[1] 지역별 에너지 균형 분석 중...")
    
    try:
        # 최신 네트워크 파일 로드
        network = get_latest_result_file()
        print(f"네트워크 파일 로드 중: {network}")
        n = pypsa.Network(network)
        
        # 발전기 출력 데이터 로드
        generator_output = get_latest_generator_output()
        print(f"발전기 출력 파일 로드 중: {generator_output}")
        generator_df = pd.read_csv(generator_output)
        
        # 부하 데이터 로드
        load_data = get_latest_load_data()
        print(f"부하 데이터 파일 로드 중: {load_data}")
        load_df = pd.read_csv(load_data)

        # 결과를 저장할 디렉토리 설정
        if output_dir is None:
            output_dir = os.path.dirname(network)
        
        # 데이터 처리를 위한 초기화
        regions = {}
        region_data = []
        
        # 모든 버스에 대한 지역 코드 추출
        for bus in n.buses.index:
            region_code = extract_region_code(bus)
            if region_code not in regions:
                regions[region_code] = {
                    'demand': 0.0,
                    'generation': 0.0,
                    'renewable_gen': 0.0
                }
        
        # 부하 데이터 처리 (지역별 수요 계산)
        for load in load_df.columns:
            bus = n.loads.at[load, 'bus']
            region_code = extract_region_code(bus)
            load_sum = load_df[load].sum()
            regions[region_code]['demand'] += load_sum
        
        # 발전기 데이터 처리 (지역별 발전량 계산)
        for gen in generator_df.columns:
            if gen not in n.generators.index:
                continue  # 발전기가 네트워크에 없으면 건너뛰기
                
            bus = n.generators.at[gen, 'bus']
            region_code = extract_region_code(bus)
            
            generation = generator_df[gen].sum()
            regions[region_code]['generation'] += generation
            
            # 재생에너지 발전량 집계
            gen_name = gen.lower()
            carrier = str(n.generators.at[gen, 'carrier']).lower()
            
            if ('pv' in gen_name or 'solar' in gen_name or 'solar' in carrier or '태양' in gen_name or
                'wt' in gen_name or 'wind' in gen_name or 'wind' in carrier or '풍력' in gen_name):
                regions[region_code]['renewable_gen'] += generation
        
        # 지역별 데이터 정리
        for region_code, data in regions.items():
            if data['demand'] > 0:
                energy_self_sufficiency = min(100, (data['generation'] / data['demand']) * 100) if data['demand'] > 0 else 0
                renewable_self_sufficiency = min(100, (data['renewable_gen'] / data['demand']) * 100) if data['demand'] > 0 else 0
            else:
                energy_self_sufficiency = 0
                renewable_self_sufficiency = 0
                
            region_data.append({
                '지역 코드': region_code,
                '지역명': get_korean_region_name(region_code),
                '수요 (MWh)': round(data['demand'], 2),
                '발전량 (MWh)': round(data['generation'], 2),
                '재생에너지 발전량 (MWh)': round(data['renewable_gen'], 2),
                '에너지 자립도 (%)': round(energy_self_sufficiency, 2),
                '재생에너지 자립도 (%)': round(renewable_self_sufficiency, 2)
            })
        
        # 데이터프레임 생성 및 정렬
        result_df = pd.DataFrame(region_data)
        result_df = result_df.sort_values('지역 코드')
        
        # 결과 저장
        output_file = os.path.join(output_dir, 'regional_energy_balance.csv')
        result_df.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f"지역별 에너지 균형 분석 결과가 '{output_file}'에 저장되었습니다.")
        
        return result_df
        
    except Exception as e:
        print(f"에러 발생: {str(e)}")
        return None

def plot_regional_energy_balance(df, output_dir=None):
    """지역별 에너지 균형 시각화"""
    if output_dir is None:
        output_dir = get_results_dir()
    
    plt.figure(figsize=(15, 10))
    
    # 서브플롯 생성
    plt.subplot(2, 1, 1)
    x = range(len(df))
    width = 0.4
    
    plt.bar(x, df['수요 (MWh)'], width=width, label='총 수요', color='skyblue')
    plt.bar([i + width for i in x], df['발전량 (MWh)'], width=width, label='총 발전량', color='lightgreen')
    
    plt.xticks([i + width/2 for i in x], df['지역명'], rotation=45, ha='right')
    plt.ylabel('에너지 (MWh)')
    plt.title('지역별 에너지 수요 및 발전량')
    plt.legend()
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    
    # 자립도 그래프
    plt.subplot(2, 1, 2)
    plt.bar(x, df['에너지 자립도 (%)'], width=width, label='에너지 자립도', color='orange')
    plt.bar([i + width for i in x], df['재생에너지 자립도 (%)'], width=width, label='재생에너지 자립도', color='green')
    
    plt.xticks([i + width/2 for i in x], df['지역명'], rotation=45, ha='right')
    plt.ylabel('자립도 (%)')
    plt.title('지역별 에너지 자립도')
    plt.legend()
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    
    plt.tight_layout()
    output_file = os.path.join(output_dir, 'regional_energy_balance.png')
    plt.savefig(output_file, dpi=300, bbox_inches='tight')
    print(f"지역별 에너지 균형 그래프가 '{output_file}'에 저장되었습니다.")
    
    # 재생에너지 비율 파이 차트
    plt.figure(figsize=(15, 10))
    
    for i, region in enumerate(df['지역명']):
        plt.subplot(3, 6, i+1)
        
        renewable = df.loc[df['지역명'] == region, '재생에너지 발전량 (MWh)'].values[0]
        conventional = df.loc[df['지역명'] == region, '발전량 (MWh)'].values[0] - renewable
        
        if conventional < 0:  # 오류 방지
            conventional = 0
        
        if renewable + conventional > 0:  # 0으로 나누기 방지
            sizes = [renewable, conventional]
            labels = ['재생에너지', '기존 발전']
            colors = ['green', 'gray']
            
            plt.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
            plt.axis('equal')
        else:
            plt.text(0.5, 0.5, "발전량 없음", horizontalalignment='center')
            
        plt.title(region, fontsize=10)
    
    plt.suptitle('지역별 재생에너지 비율', fontsize=16)
    plt.tight_layout()
    
    # 결과 저장
    output_file = os.path.join(output_dir, 'regional_renewable_ratio.png')
    plt.savefig(output_file, dpi=300, bbox_inches='tight')
    print(f"지역별 재생에너지 비율 그래프가 '{output_file}'에 저장되었습니다.")

def analyze_transmission_flow(output_dir=None):
    """송전망 흐름 분석"""
    try:
        if output_dir is None:
            output_dir = get_results_dir()
            
        # 최신 결과 파일 로드
        network_file = get_latest_result_file()
        print(f"네트워크 파일 로드 중: {network_file}")
        network = pypsa.Network(network_file)
        
        # 라인 사용률 데이터 로드
        line_usage_file = get_latest_line_usage()
        print(f"라인 사용률 파일 로드 중: {line_usage_file}")
        line_usage = pd.read_csv(line_usage_file, index_col=0)
        
        # 라인별 총 전력 흐름 계산
        line_flow_data = []
        
        for line in network.lines.index:
            if line in line_usage.columns:
                # 절대값의 합 계산 (양방향 흐름의 총량)
                total_flow = line_usage[line].abs().sum()
                
                bus0 = network.lines.at[line, 'bus0']
                bus1 = network.lines.at[line, 'bus1']
                region0 = extract_region_code(bus0)
                region1 = extract_region_code(bus1)
                
                line_flow_data.append({
                    '선로명': line,
                    '시작 버스': bus0,
                    '종료 버스': bus1,
                    '시작 지역': get_korean_region_name(region0),
                    '종료 지역': get_korean_region_name(region1),
                    '총 전력 흐름 (MWh)': round(total_flow, 2)
                })
        
        # 데이터프레임 생성 및 정렬
        df = pd.DataFrame(line_flow_data)
        df = df.sort_values('총 전력 흐름 (MWh)', ascending=False)
        
        # 결과 저장
        output_file = os.path.join(output_dir, 'transmission_flow.csv')
        df.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f"송전망 흐름 분석이 '{output_file}'에 저장되었습니다.")
        
        # 테이블 형태로 결과 출력
        print("\n송전망 흐름 분석 결과 (상위 10개):")
        print(df.head(10).to_string(index=False))
        
        # 시각화 - 명시적으로 output_dir 전달
        create_transmission_flow_map(df, network, output_dir)
        
        # 각 지역의 대표 좌표
        region_coords = {
            '서울특별시': [37.5665, 126.9780],
            '부산광역시': [35.1796, 129.0756],
            '인천광역시': [37.4563, 126.7052],
            '대구광역시': [35.8714, 128.6014],
            '광주광역시': [35.1595, 126.8526],
            '대전광역시': [36.3504, 127.3845],
            '울산광역시': [35.5384, 129.3114],
            '세종특별자치시': [36.4800, 127.2890],
            '경기도': [37.4138, 127.5183],
            '강원도': [37.8228, 128.1555],
            '충청북도': [36.6358, 127.4915],
            '충청남도': [36.6588, 126.6728],
            '전라북도': [35.7175, 127.1530],
            '전라남도': [34.8679, 126.9910],
            '경상북도': [36.4919, 128.8889],
            '경상남도': [35.4606, 128.2132],
            '제주특별자치도': [33.4890, 126.4983]
        }
        
        # 명시적으로 output_dir 전달
        create_network_graph(df, region_coords, output_dir)
        
        return df
        
    except Exception as e:
        print(f"에러 발생: {str(e)}")
        traceback.print_exc()
        return None

def create_transmission_flow_map(transmission_df, network, output_dir=None):
    """송전망 흐름 지도 생성"""
    try:
        if output_dir is None:
            output_dir = get_results_dir()
            
        # 한국 지도의 중심점
        center = [36.0, 128.0]
        
        # 각 지역의 대표 좌표 (실제 좌표로 업데이트 필요)
        region_coords = {
            '서울특별시': [37.5665, 126.9780],
            '부산광역시': [35.1796, 129.0756],
            '인천광역시': [37.4563, 126.7052],
            '대구광역시': [35.8714, 128.6014],
            '광주광역시': [35.1595, 126.8526],
            '대전광역시': [36.3504, 127.3845],
            '울산광역시': [35.5384, 129.3114],
            '세종특별자치시': [36.4800, 127.2890],
            '경기도': [37.4138, 127.5183],
            '강원도': [37.8228, 128.1555],
            '충청북도': [36.6358, 127.4915],
            '충청남도': [36.6588, 126.6728],
            '전라북도': [35.7175, 127.1530],
            '전라남도': [34.8679, 126.9910],
            '경상북도': [36.4919, 128.8889],
            '경상남도': [35.4606, 128.2132],
            '제주특별자치도': [33.4890, 126.4983]
        }
        
        # 지도 생성
        m = folium.Map(location=center, zoom_start=7, tiles='CartoDB positron')
        
        # 지역 마커 추가
        for region, coords in region_coords.items():
            folium.Marker(
                location=coords,
                popup=region,
                tooltip=region,
                icon=folium.Icon(color='blue', icon='info-sign')
            ).add_to(m)
        
        # 송전선 추가
        flow_values = [row['총 전력 흐름 (MWh)'] for _, row in transmission_df.iterrows()]
        max_flow = max(flow_values) if flow_values else 1
        
        for _, row in transmission_df.iterrows():
            start_region = row['시작 지역']
            end_region = row['종료 지역']
            flow = row['총 전력 흐름 (MWh)']
            
            # 지역 좌표가 있는 경우에만 선 추가
            if start_region in region_coords and end_region in region_coords:
                # 선 두께 계산 (최대값 대비 비율)
                weight = 1 + 9 * (flow / max_flow)  # 1~10 범위
                color = 'red' if flow > max_flow * 0.7 else 'orange' if flow > max_flow * 0.4 else 'blue'
                
                folium.PolyLine(
                    locations=[region_coords[start_region], region_coords[end_region]],
                    color=color,
                    weight=weight,
                    opacity=0.8,
                    popup=f"{start_region} → {end_region}: {flow:.2f} MWh"
                ).add_to(m)
        
        # 범례 추가
        legend_html = '''
             <div style="position: fixed; 
                         bottom: 50px; left: 50px; width: 230px; height: 160px; 
                         border:2px solid grey; z-index:9999; font-size:14px;
                         background-color: white; padding: 10px;
                         opacity: 0.8">
             <div style="margin-bottom: 5px;"><b>송전량 범례</b></div>
             <div style="display: flex; align-items: center; margin-bottom: 5px;">
               <div style="background-color: blue; width: 30px; height: 5px; margin-right: 10px;"></div>
               <div>낮음 (&lt; 40%)</div>
             </div>
             <div style="display: flex; align-items: center; margin-bottom: 5px;">
               <div style="background-color: orange; width: 30px; height: 5px; margin-right: 10px;"></div>
               <div>중간 (40-70%)</div>
             </div>
             <div style="display: flex; align-items: center;">
               <div style="background-color: red; width: 30px; height: 5px; margin-right: 10px;"></div>
               <div>높음 (&gt; 70%)</div>
             </div>
             <div style="margin-top: 10px; font-size: 12px;">* 비율은 최대 송전량 대비 %</div>
             </div>
             '''
        m.get_root().html.add_child(folium.Element(legend_html))
        
        # 지도 저장 - output_dir 폴더에 명시적으로 저장
        output_file = os.path.join(output_dir, 'transmission_flow_map.html')
        m.save(output_file)
        print(f"송전망 흐름 지도가 '{output_file}'에 저장되었습니다.")
        
        return m
        
    except Exception as e:
        print(f"송전망 흐름 지도 생성 중 오류 발생: {str(e)}")
        traceback.print_exc()
        return None

def create_network_graph(transmission_df, region_coords, output_dir=None):
    """송전망 그래프 시각화"""
    try:
        if output_dir is None:
            output_dir = get_results_dir()
            
        # 그래프 생성
        G = nx.Graph()
        
        # 지역을 노드로 추가
        for region, coords in region_coords.items():
            G.add_node(region, pos=coords)
        
        # 송전선을 엣지로 추가
        for _, row in transmission_df.iterrows():
            start_region = row['시작 지역']
            end_region = row['종료 지역']
            flow = row['총 전력 흐름 (MWh)']
            
            # 두 지역 모두 그래프에 있는 경우에만 엣지 추가
            if start_region in G.nodes and end_region in G.nodes:
                G.add_edge(start_region, end_region, weight=flow)
        
        # 그래프 그리기
        plt.figure(figsize=(18, 15))
        
        # 노드 위치 설정
        pos = nx.get_node_attributes(G, 'pos')
        
        # 엣지 가중치 추출
        weights = [G[u][v]['weight'] for u, v in G.edges()]
        if not weights:  # 빈 리스트일 경우 예외 처리
            print("그래프 생성 불가: 전력 흐름 데이터가 없습니다.")
            return
            
        max_weight = max(weights)
        normalized_weights = [w / max_weight * 10 for w in weights]
        
        # 엣지 그리기 (선 두께가 가중치에 비례)
        nx.draw_networkx_edges(G, pos, width=normalized_weights, edge_color='gray', alpha=0.7)
        
        # 노드 그리기
        nx.draw_networkx_nodes(G, pos, node_size=500, node_color='skyblue', alpha=0.8)
        
        # 라벨 그리기 (KoPubWorld Dotum Medium 폰트로)
        try:
            font_path = 'KoPubWorld Dotum Medium.ttf'
            if os.path.exists(font_path):
                # 노드 레이블 직접 추가 (폰트 속성 명시)
                for node, (x, y) in pos.items():
                    plt.text(x, y, node, fontsize=12, ha='center', va='center',
                            bbox=dict(facecolor='white', alpha=0.6, edgecolor='none'),
                            fontfamily='KoPubWorld Dotum Medium')
            else:
                # 기본 폰트로 노드 레이블 추가
                nx.draw_networkx_labels(G, pos, font_size=12, font_family='sans-serif')
        except Exception as font_e:
            print(f"폰트 설정 중 오류 발생: {str(font_e)}")
            # 기본 방식으로 레이블 추가
            nx.draw_networkx_labels(G, pos, font_size=12)
        
        # 엣지 레이블 추가
        edge_labels = {(u, v): f"{G[u][v]['weight']:.0f} MWh" for u, v in G.edges()}
        try:
            # 폰트 속성 명시하여 엣지 레이블 직접 추가
            for (u, v), label in edge_labels.items():
                x = (pos[u][0] + pos[v][0]) / 2
                y = (pos[u][1] + pos[v][1]) / 2
                plt.text(x, y, label, fontsize=8, ha='center', va='center',
                        bbox=dict(facecolor='white', alpha=0.7, edgecolor='none'),
                        fontfamily='KoPubWorld Dotum Medium')
        except Exception as edge_e:
            print(f"엣지 레이블 설정 중 오류 발생: {str(edge_e)}")
            # 기본 방식으로 엣지 레이블 추가
            nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, font_size=8)
        
        # 타이틀 설정
        plt.title('지역간 송전망 흐름도', fontsize=20)
        plt.axis('off')
        plt.tight_layout()
        
        # 그래프 저장 - output_dir 폴더에 명시적으로 저장
        output_file = os.path.join(output_dir, 'transmission_network_graph.png')
        plt.savefig(output_file, dpi=300, bbox_inches='tight')
        print(f"송전망 네트워크 그래프가 '{output_file}'에 저장되었습니다.")
    
    except Exception as e:
        print(f"네트워크 그래프 생성 중 오류 발생: {str(e)}")
        traceback.print_exc()

def analyze_generators_by_region_and_carrier(output_dir=None):
    """지역별 발전원별 발전량 분석"""
    try:
        if output_dir is None:
            output_dir = get_results_dir()
            
        # 최신 결과 파일 로드
        network_file = get_latest_result_file()
        print(f"발전기 분석: 네트워크 파일 로드 중: {network_file}")
        network = pypsa.Network(network_file)
        
        # 발전기 출력 데이터 로드
        gen_output_file = get_latest_generator_output()
        print(f"발전기 분석: 발전기 출력 파일 로드 중: {gen_output_file}")
        gen_output = pd.read_csv(gen_output_file, index_col=0)
        
        # 지역별 발전원별 발전량 분석
        region_carrier_data = []
        
        # 발전기별 총 발전량 계산
        gen_total = {col: gen_output[col].sum() for col in gen_output.columns if col in network.generators.index}
        
        # 각 발전기에 대해 지역과 에너지원 확인
        for gen_name, total_output in gen_total.items():
            if gen_name not in network.generators.index:
                continue
                
            bus = network.generators.at[gen_name, 'bus']
            region_code = extract_region_code(bus)
            region_name = get_korean_region_name(region_code)
            
            # 에너지원(carrier) 정보 가져오기
            carrier = str(network.generators.at[gen_name, 'carrier']).lower()
            
            # 발전원 표준화 - 모든 발전기에 대해 수행
            gen_lower = gen_name.lower()
            
            # 발전기 이름에서 발전원 유형 추론
            if 'pv' in gen_lower or 'solar' in gen_lower or '태양' in gen_lower:
                carrier = 'solar'
            elif 'wind' in gen_lower or 'wt' in gen_lower or '풍력' in gen_lower:
                carrier = 'wind'
            elif 'hydro' in gen_lower or '수력' in gen_lower:
                carrier = 'hydro'
            elif 'nuclear' in gen_lower or '원자력' in gen_lower:
                carrier = 'nuclear'
            elif 'coal' in gen_lower or '석탄' in gen_lower:
                carrier = 'coal'
            elif 'gas' in gen_lower or 'ng' in gen_lower or 'lng' in gen_lower or '가스' in gen_lower:
                carrier = 'gas'
            elif 'oil' in gen_lower or '석유' in gen_lower:
                carrier = 'oil'
            elif 'biomass' in gen_lower or '바이오' in gen_lower:
                carrier = 'biomass'
            # carrier 값이 유효한 경우 (electricity, hydrogen이 아니고 이미 구체적인 발전원인 경우)
            elif carrier not in ['nan', '', 'electricity', 'hydrogen', 'electric'] and carrier in ['solar', 'wind', 'hydro', 'nuclear', 'coal', 'gas', 'oil', 'biomass']:
                # 이미 구체적인 발전원이므로 그대로 사용
                pass
            else:
                # 기본값
                carrier = 'other'
            
            # 한글 에너지원 이름
            carrier_korean = {
                'solar': '태양광',
                'wind': '풍력',
                'hydro': '수력',
                'nuclear': '원자력',
                'coal': '석탄',
                'gas': 'LNG',
                'oil': '석유',
                'biomass': '바이오매스',
                'other': '기타'
            }.get(carrier, carrier)
            
            region_carrier_data.append({
                '지역': region_name,
                '에너지원': carrier_korean,
                '발전량_MWh': total_output,
                '발전기': gen_name
            })
        
        # 데이터프레임 생성
        region_carrier_df = pd.DataFrame(region_carrier_data)
        
        # 1. 지역별 발전원별 발전량
        region_carrier_pivot = region_carrier_df.pivot_table(
            index='지역', 
            columns='에너지원', 
            values='발전량_MWh', 
            aggfunc='sum',
            fill_value=0
        ).reset_index()
        
        # 전체 합계 추가
        region_carrier_pivot['총발전량'] = region_carrier_pivot.drop('지역', axis=1).sum(axis=1)
        
        # CSV 저장
        output_file = os.path.join(output_dir, f'optimization_result_{get_latest_timestamp()}_지역별_발전원별_발전량.csv')
        region_carrier_pivot.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f"지역별 발전원별 발전량이 '{output_file}'에 저장되었습니다.")
        
        # 2. 지역별 발전량
        region_generation = region_carrier_df.groupby('지역')['발전량_MWh'].sum().reset_index()
        total_generation = region_generation['발전량_MWh'].sum()
        region_generation['비율'] = region_generation['발전량_MWh'] / total_generation * 100
        region_generation = region_generation.sort_values('발전량_MWh', ascending=False)
        
        output_file = os.path.join(output_dir, f'optimization_result_{get_latest_timestamp()}_지역별_발전량.csv')
        region_generation.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f"지역별 발전량이 '{output_file}'에 저장되었습니다.")
        
        # 3. 발전원별 발전량
        carrier_generation = region_carrier_df.groupby('에너지원')['발전량_MWh'].sum().reset_index()
        carrier_generation['비율'] = carrier_generation['발전량_MWh'] / total_generation * 100
        carrier_generation = carrier_generation.sort_values('발전량_MWh', ascending=False)
        
        output_file = os.path.join(output_dir, f'optimization_result_{get_latest_timestamp()}_발전원별_발전량.csv')
        carrier_generation.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f"발전원별 발전량이 '{output_file}'에 저장되었습니다.")
        
        # 4. 상위 발전기 발전량
        generator_output = region_carrier_df[['발전기', '지역', '에너지원', '발전량_MWh']]
        generator_output['비율'] = generator_output['발전량_MWh'] / total_generation * 100
        generator_output = generator_output.sort_values('발전량_MWh', ascending=False)
        
        # 상위 100개 발전기만 저장
        top_generators = generator_output.head(100)
        
        output_file = os.path.join(output_dir, f'optimization_result_{get_latest_timestamp()}_상위발전기_발전량.csv')
        top_generators.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f"상위발전기 발전량이 '{output_file}'에 저장되었습니다.")
        
        return {
            'region_carrier': region_carrier_pivot,
            'region': region_generation,
            'carrier': carrier_generation,
            'top_generators': top_generators
        }
        
    except Exception as e:
        print(f"발전기 분석 중 오류 발생: {str(e)}")
        traceback.print_exc()
        return None

def main(output_dir=None):
    """메인 함수"""
    print("=" * 80)
    print("지역별 에너지 분석 및 송전망 흐름 시각화")
    print("=" * 80)
    
    # 결과 폴더 확인
    if output_dir is None:
        try:
            output_dir = get_results_dir()
            print(f"결과 파일은 '{output_dir}' 폴더에 저장됩니다.")
        except Exception as e:
            print(f"결과 폴더 확인 중 오류 발생: {str(e)}")
            output_dir = RESULTS_DIR
    else:
        print(f"결과 파일은 '{output_dir}' 폴더에 저장됩니다.")
    
    # 지역별 에너지 균형 분석
    print("\n[1] 지역별 에너지 균형 분석 중...")
    regional_df = analyze_regional_energy_balance(output_dir)
    plot_regional_energy_balance(regional_df, output_dir=output_dir)
    
    # 송전망 흐름 분석
    print("\n[2] 송전망 흐름 분석 중...")
    transmission_df = analyze_transmission_flow(output_dir=output_dir)
    
    # 발전기 상세 분석 (지역별, 발전원별)
    print("\n[3] 발전기 상세 분석 중...")
    generator_analysis = analyze_generators_by_region_and_carrier(output_dir=output_dir)
    
    print("\n분석 완료!")
    print("=" * 80)
    print("생성된 파일 목록:")
    print(f"1. {os.path.join(output_dir, 'regional_energy_balance.csv')} - 지역별 에너지 균형 데이터")
    print(f"2. {os.path.join(output_dir, 'regional_energy_balance.png')} - 지역별 에너지 균형 그래프")
    print(f"3. {os.path.join(output_dir, 'regional_renewable_ratio.png')} - 지역별 재생에너지 비율 그래프")
    print(f"4. {os.path.join(output_dir, 'transmission_flow.csv')} - 송전망 흐름 데이터")
    print(f"5. {os.path.join(output_dir, 'transmission_flow_map.html')} - 송전망 흐름 지도 (인터랙티브)")
    print(f"6. {os.path.join(output_dir, 'transmission_network_graph.png')} - 송전망 네트워크 그래프")
    print(f"7. {os.path.join(output_dir, f'optimization_result_{get_latest_timestamp()}_지역별_발전원별_발전량.csv')} - 지역별 발전원별 발전량")
    print(f"8. {os.path.join(output_dir, f'optimization_result_{get_latest_timestamp()}_지역별_발전량.csv')} - 지역별 발전량")
    print(f"9. {os.path.join(output_dir, f'optimization_result_{get_latest_timestamp()}_발전원별_발전량.csv')} - 발전원별 발전량")
    print(f"10. {os.path.join(output_dir, f'optimization_result_{get_latest_timestamp()}_상위발전기_발전량.csv')} - 상위발전기 발전량")
    print("=" * 80)

if __name__ == "__main__":
    main() 