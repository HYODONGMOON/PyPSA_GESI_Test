import pandas as pd
import numpy as np
import networkx as nx
from datetime import datetime
import shutil

def implement_distance_based_transmission():
    print("=== 거리 기반 송전 손실률 모델링 구현 ===")
    
    try:
        # 백업 생성
        backup_file = f"integrated_input_data_before_distance_modeling_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        shutil.copy2('integrated_input_data.xlsx', backup_file)
        print(f"백업 파일 생성: {backup_file}")
        
        # 데이터 로드
        input_data = {}
        with pd.ExcelFile('integrated_input_data.xlsx') as xls:
            for sheet in xls.sheet_names:
                input_data[sheet] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet)
        
        lines_df = input_data['lines'].copy()
        print(f"기존 송전선로: {len(lines_df)}개")
        
        # 1. 지역간 최단 거리 계산
        print("\n1. 지역간 최단 거리 매트릭스 생성...")
        shortest_distances = calculate_shortest_distances(lines_df)
        
        # 2. 거리 기반 저항 및 효율 계산
        print("\n2. 거리 기반 송전 모델 적용...")
        enhanced_lines_df = apply_distance_based_model(lines_df, shortest_distances)
        
        # 3. 결과 저장
        input_data['lines'] = enhanced_lines_df
        
        with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
            for sheet_name, df in input_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n✅ 거리 기반 송전 모델링 완료!")
        print(f"향상된 송전선로 모델이 적용되었습니다.")
        
        # 4. 변경사항 요약
        print(f"\n=== 적용된 모델링 요약 ===")
        print(f"- 기준 저항률: 0.048 Ω/km (AC 345kV 기준)")
        print(f"- 기본 송전 손실률: 2% per 100km")
        print(f"- 우회 경로 페널티: 거리 비례 (최대 2배)")
        print(f"- 허브 지역 보너스: CBD, GND, JBD 15% 효율 증가")
        
        return True
        
    except Exception as e:
        print(f"구현 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

def calculate_shortest_distances(lines_df):
    """지역간 최단 거리 매트릭스 계산"""
    print("   지역간 네트워크 그래프 생성...")
    
    # 네트워크 그래프 생성
    G = nx.Graph()
    
    for _, line in lines_df.iterrows():
        bus0 = line['bus0']
        bus1 = line['bus1']
        length = line.get('length', 100)  # 기본값 100km
        
        # 지역 코드 추출
        region0 = bus0.split('_')[0] if '_' in bus0 else bus0[:3]
        region1 = bus1.split('_')[0] if '_' in bus1 else bus1[:3]
        
        if region0 != region1:  # 지역간 연결만
            if G.has_edge(region0, region1):
                # 기존 연결이 있으면 최소 거리 사용
                G[region0][region1]['length'] = min(G[region0][region1]['length'], length)
            else:
                G.add_edge(region0, region1, length=length)
    
    # 모든 지역 쌍의 최단 거리 계산
    print("   최단 경로 계산 중...")
    shortest_distances = {}
    
    for source in G.nodes():
        shortest_distances[source] = nx.single_source_dijkstra_path_length(G, source, weight='length')
    
    print(f"   완료: {len(G.nodes())}개 지역, {len(G.edges())}개 연결")
    return shortest_distances

def apply_distance_based_model(lines_df, shortest_distances):
    """거리 기반 송전 모델 적용"""
    
    enhanced_lines = lines_df.copy()
    
    # 모델 파라미터
    BASE_RESISTANCE = 0.048  # Ω/km (AC 345kV 기준)
    BASE_LOSS_RATE = 0.02    # 2% per 100km
    DETOUR_PENALTY_FACTOR = 0.3  # 우회 시 30% 추가 페널티
    HUB_EFFICIENCY_BONUS = 0.15  # 허브 지역 15% 효율 증가
    
    # 허브 지역 정의 (중개 중심성이 높은 지역)
    HUB_REGIONS = ['CBD', 'GND', 'JBD', 'GGD', 'GBD']
    
    print("   개별 송전선로 모델링 중...")
    
    for idx, line in enhanced_lines.iterrows():
        bus0 = line['bus0']
        bus1 = line['bus1']
        actual_length = line.get('length', 100)
        
        # 지역 코드 추출
        region0 = bus0.split('_')[0] if '_' in bus0 else bus0[:3]
        region1 = bus1.split('_')[0] if '_' in bus1 else bus1[:3]
        
        if region0 == region1:  # 지역 내 연결은 그대로
            continue
            
        # 최단 거리 대비 실제 거리 비율
        try:
            shortest_distance = shortest_distances[region0].get(region1, actual_length)
            detour_ratio = actual_length / shortest_distance if shortest_distance > 0 else 1.0
        except:
            detour_ratio = 1.0
        
        # 1. 거리 기반 저항 계산
        base_resistance = BASE_RESISTANCE * actual_length
        
        # 우회 페널티 적용
        if detour_ratio > 1.2:  # 20% 이상 우회 시 페널티
            detour_penalty = 1 + DETOUR_PENALTY_FACTOR * (detour_ratio - 1)
            adjusted_resistance = base_resistance * detour_penalty
        else:
            adjusted_resistance = base_resistance
        
        # 2. 허브 지역 보너스
        if region0 in HUB_REGIONS or region1 in HUB_REGIONS:
            adjusted_resistance *= (1 - HUB_EFFICIENCY_BONUS)
        
        # 3. 송전 손실률 기반 실효 용량 계산
        transmission_loss = BASE_LOSS_RATE * (actual_length / 100)
        efficiency = 1 - transmission_loss
        
        # 우회 경로 효율 페널티
        if detour_ratio > 1.5:  # 50% 이상 우회 시 추가 효율 감소
            efficiency *= (1 - 0.1 * (detour_ratio - 1.5))
        
        # 허브 지역 효율 보너스
        if region0 in HUB_REGIONS or region1 in HUB_REGIONS:
            efficiency *= (1 + HUB_EFFICIENCY_BONUS)
        
        efficiency = max(0.5, min(0.98, efficiency))  # 50%-98% 범위로 제한
        
        # 4. 실효 용량 조정
        original_capacity = line['s_nom']
        effective_capacity = original_capacity * efficiency
        
        # 결과 적용
        enhanced_lines.at[idx, 'r'] = adjusted_resistance
        enhanced_lines.at[idx, 's_nom'] = effective_capacity
        
        # 디버그 정보 (주요 변경사항만 출력)
        if detour_ratio > 1.3 or region0 in HUB_REGIONS or region1 in HUB_REGIONS:
            print(f"      {line['name']}: {region0}-{region1}")
            print(f"        거리 비율: {detour_ratio:.2f}x")
            print(f"        효율: {efficiency:.1%}")
            print(f"        용량: {original_capacity:.0f} → {effective_capacity:.0f} MW")
    
    # 5. 추가 우회 경로 생성 (병목 지역 해소용)
    print("\n   병목 해소용 가상 우회 경로 추가...")
    additional_lines = create_virtual_detour_paths(enhanced_lines, shortest_distances)
    
    if additional_lines is not None and not additional_lines.empty:
        enhanced_lines = pd.concat([enhanced_lines, additional_lines], ignore_index=True)
        print(f"   추가된 가상 우회 경로: {len(additional_lines)}개")
    
    print(f"\n   완료: {len(enhanced_lines)}개 송전선로 (기존 {len(lines_df)}개 + 추가 {len(enhanced_lines) - len(lines_df)}개)")
    
    return enhanced_lines

def create_virtual_detour_paths(lines_df, shortest_distances):
    """병목 지역 해소를 위한 가상 우회 경로 생성"""
    
    # 병목이 심한 연결들
    bottleneck_connections = [
        ("DGU", "GND", 1000),  # 1000MW 추가 용량
        ("DJN", "SJN", 800),   # 800MW 추가 용량
        ("JJD", "JND", 600),   # 600MW 추가 용량 (제주 연결 강화)
        ("GND", "JND", 400),   # 400MW 추가 용량
        ("CND", "GGD", 600)    # 600MW 추가 용량
    ]
    
    additional_lines = []
    
    for region1, region2, additional_capacity in bottleneck_connections:
        # 기존 연결 찾기
        existing_line = None
        for _, line in lines_df.iterrows():
            bus0_region = line['bus0'].split('_')[0] if '_' in line['bus0'] else line['bus0'][:3]
            bus1_region = line['bus1'].split('_')[0] if '_' in line['bus1'] else line['bus1'][:3]
            
            if (bus0_region == region1 and bus1_region == region2) or (bus0_region == region2 and bus1_region == region1):
                existing_line = line
                break
        
        if existing_line is not None:
            # 가상 우회 경로 생성
            new_line = existing_line.copy()
            new_line['name'] = f"{region1}_{region2}_가상우회_{additional_capacity}MW"
            new_line['s_nom'] = additional_capacity
            
            # 우회 경로는 더 높은 저항 (비용)을 가짐
            new_line['r'] = existing_line.get('r', 0.1) * 1.8  # 80% 높은 저항
            
            additional_lines.append(new_line)
    
    return pd.DataFrame(additional_lines) if additional_lines else None

if __name__ == "__main__":
    implement_distance_based_transmission() 