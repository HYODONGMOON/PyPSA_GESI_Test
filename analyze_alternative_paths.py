import pandas as pd
import numpy as np
import networkx as nx

def analyze_alternative_paths():
    print("=== 우회 송전 경로 분석 ===")
    
    try:
        # 송전선로 데이터 로드
        lines_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='lines')
        
        # 네트워크 그래프 생성
        G = nx.Graph()
        
        for _, line in lines_df.iterrows():
            bus0 = line['bus0']
            bus1 = line['bus1']
            capacity = line['s_nom']
            length = line.get('length', 100)  # 기본값 100km
            resistance = line.get('r', 0.1)   # 기본값 0.1
            
            # 지역 코드 추출
            region0 = bus0.split('_')[0] if '_' in bus0 else bus0[:3]
            region1 = bus1.split('_')[0] if '_' in bus1 else bus1[:3]
            
            if region0 != region1:  # 지역간 연결만
                # 가중치: 저항 고려 (저항이 낮을수록 우선)
                weight = resistance + (length / 1000)  # 거리와 저항 조합
                
                if G.has_edge(region0, region1):
                    # 기존 연결이 있으면 용량 합산, 최소 가중치 사용
                    G[region0][region1]['capacity'] += capacity
                    G[region0][region1]['weight'] = min(G[region0][region1]['weight'], weight)
                else:
                    G.add_edge(region0, region1, capacity=capacity, weight=weight)
        
        print(f"지역: {len(G.nodes)}개")
        print(f"지역간 연결: {len(G.edges)}개")
        
        # 문제 지역들의 우회 경로 분석
        problematic_connections = [
            ("DGU", "GND"),
            ("DJN", "SJN"), 
            ("JJD", "JND"),
            ("GND", "JND"),
            ("CND", "GGD")
        ]
        
        print("\n=== 우회 경로 분석 ===")
        
        for source, target in problematic_connections:
            print(f"\n🔍 {source} → {target} 우회 경로 분석:")
            
            if not G.has_node(source) or not G.has_node(target):
                print(f"   ⚠️ {source} 또는 {target} 지역이 네트워크에 없음")
                continue
            
            # 직접 연결 용량
            direct_capacity = 0
            if G.has_edge(source, target):
                direct_capacity = G[source][target]['capacity']
                print(f"   직접 연결: {direct_capacity:,.0f} MW")
            else:
                print(f"   직접 연결: 없음")
            
            # 모든 가능한 경로 찾기 (최대 4개 홉까지)
            try:
                all_paths = list(nx.all_simple_paths(G, source, target, cutoff=4))
                print(f"   발견된 경로: {len(all_paths)}개")
                
                # 경로별 분석
                path_analysis = []
                
                for path in all_paths[:10]:  # 상위 10개 경로만 분석
                    # 경로의 최소 용량 (병목)
                    min_capacity = float('inf')
                    total_weight = 0
                    
                    for i in range(len(path) - 1):
                        edge_data = G[path[i]][path[i+1]]
                        min_capacity = min(min_capacity, edge_data['capacity'])
                        total_weight += edge_data['weight']
                    
                    path_analysis.append({
                        'path': ' → '.join(path),
                        'hops': len(path) - 1,
                        'bottleneck_capacity': min_capacity,
                        'total_weight': total_weight
                    })
                
                # 용량별 정렬
                path_analysis.sort(key=lambda x: x['bottleneck_capacity'], reverse=True)
                
                print(f"\n   📊 주요 우회 경로 (병목 용량 순):")
                total_alternative_capacity = 0
                
                for i, path_info in enumerate(path_analysis[:5]):
                    if path_info['hops'] == 1:  # 직접 연결
                        print(f"   {i+1}. {path_info['path']} (직접)")
                    else:  # 우회 경로
                        print(f"   {i+1}. {path_info['path']}")
                    
                    print(f"      병목 용량: {path_info['bottleneck_capacity']:,.0f} MW")
                    print(f"      홉 수: {path_info['hops']}")
                    print(f"      가중치: {path_info['total_weight']:.2f}")
                    
                    if path_info['hops'] > 1:  # 우회 경로만
                        total_alternative_capacity += path_info['bottleneck_capacity']
                    print()
                
                # 우회 송전 가능성 평가
                print(f"   💡 우회 송전 분석:")
                print(f"   - 직접 연결 용량: {direct_capacity:,.0f} MW")
                print(f"   - 주요 우회 경로 용량: {total_alternative_capacity:,.0f} MW")
                
                if total_alternative_capacity > direct_capacity * 0.5:
                    print(f"   ✅ 우회 경로 활용 가능 (직접 연결의 50% 이상)")
                else:
                    print(f"   ❌ 우회 경로 제한적 (직접 연결 대비 부족)")
                
            except nx.NetworkXNoPath:
                print(f"   ❌ {source}와 {target} 사이에 연결된 경로가 없음")
            except Exception as e:
                print(f"   ⚠️ 경로 분석 중 오류: {e}")
        
        # 전체 네트워크 연결성 분석
        print(f"\n=== 전체 네트워크 연결성 ===")
        
        if nx.is_connected(G):
            print("✅ 모든 지역이 연결되어 있음")
            
            # 네트워크 중심성 분석
            centrality = nx.betweenness_centrality(G, weight='weight')
            print(f"\n📍 주요 허브 지역 (중개 중심성 순):")
            
            sorted_centrality = sorted(centrality.items(), key=lambda x: x[1], reverse=True)
            for region, score in sorted_centrality[:10]:
                print(f"   {region}: {score:.3f}")
            
            # 가장 중요한 연결 (bridge)
            bridges = list(nx.bridges(G))
            if bridges:
                print(f"\n🌉 핵심 연결 (끊어지면 네트워크 분리):")
                for bridge in bridges[:5]:
                    capacity = G[bridge[0]][bridge[1]]['capacity']
                    print(f"   {bridge[0]} ↔ {bridge[1]}: {capacity:,.0f} MW")
            else:
                print(f"\n✅ 단일 연결 끊어짐으로 네트워크가 분리되지 않음")
        else:
            print("❌ 일부 지역이 분리되어 있음")
            components = list(nx.connected_components(G))
            print(f"   연결된 그룹: {len(components)}개")
            for i, component in enumerate(components):
                print(f"   그룹 {i+1}: {list(component)}")
        
    except Exception as e:
        print(f"분석 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    analyze_alternative_paths() 