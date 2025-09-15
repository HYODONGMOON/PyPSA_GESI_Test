import pandas as pd
import numpy as np

def reanalyze_transmission_properly():
    print("=== 송전선로 올바른 재분석 ===")
    
    try:
        # 1. 성공한 케이스의 시간별 전력 흐름 로드
        print("1. 시간별 전력 흐름 데이터 로드 중...")
        line_usage_df = pd.read_csv('results/20250915_151959/optimization_result_20250915_151959_line_usage.csv', index_col=0)
        
        # 2. 현재 실제 송전선로 용량 로드
        lines_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='lines')
        
        # 3. 지역간 연결별로 그룹화하여 총 송전량 계산
        print("\n2. 지역간 실제 송전량 분석 중...")
        
        # 지역간 연결 그룹 생성
        regional_flows = {}
        
        for time_idx in line_usage_df.index[:100]:  # 처음 100개 시간만 분석 (속도 향상)
            regional_flows[time_idx] = {}
            
            for line_name in line_usage_df.columns:
                # 송전선로 정보 찾기
                line_info = lines_df[lines_df['name'] == line_name]
                if line_info.empty:
                    continue
                    
                bus0 = line_info.iloc[0]['bus0']
                bus1 = line_info.iloc[0]['bus1']
                
                # 지역 코드 추출
                region0 = bus0.split('_')[0] if '_' in bus0 else bus0[:3]
                region1 = bus1.split('_')[0] if '_' in bus1 else bus1[:3]
                
                if region0 != region1:
                    # 정규화된 연결명
                    connection = f"{region0}-{region1}" if region0 < region1 else f"{region1}-{region0}"
                    
                    if connection not in regional_flows[time_idx]:
                        regional_flows[time_idx][connection] = 0
                    
                    # 해당 시간의 전력 흐름 추가 (방향 고려)
                    flow_value = line_usage_df.loc[time_idx, line_name]
                    if region0 > region1:  # 방향 보정
                        flow_value = -flow_value
                    
                    regional_flows[time_idx][connection] += flow_value
        
        # 4. 지역간 최대 송전량 분석
        print("\n3. 지역간 최대 송전량 vs 총 용량 비교...")
        
        regional_max_flows = {}
        regional_total_capacity = {}
        
        # 총 용량 계산
        for _, line in lines_df.iterrows():
            bus0 = line['bus0']
            bus1 = line['bus1']
            region0 = bus0.split('_')[0] if '_' in bus0 else bus0[:3]
            region1 = bus1.split('_')[0] if '_' in bus1 else bus1[:3]
            
            if region0 != region1:
                connection = f"{region0}-{region1}" if region0 < region1 else f"{region1}-{region0}"
                
                if connection not in regional_total_capacity:
                    regional_total_capacity[connection] = 0
                
                regional_total_capacity[connection] += line['s_nom']
        
        # 최대 흐름 계산
        for time_idx in regional_flows:
            for connection, flow in regional_flows[time_idx].items():
                if connection not in regional_max_flows:
                    regional_max_flows[connection] = {'max_abs_flow': 0, 'time': None}
                
                abs_flow = abs(flow)
                if abs_flow > regional_max_flows[connection]['max_abs_flow']:
                    regional_max_flows[connection]['max_abs_flow'] = abs_flow
                    regional_max_flows[connection]['time'] = time_idx
        
        # 5. 용량 부족 지역 확인
        print("\n=== 지역간 송전 용량 vs 최대 필요량 비교 ===")
        
        capacity_issues = []
        
        for connection in regional_total_capacity:
            total_capacity = regional_total_capacity[connection]
            max_flow = regional_max_flows.get(connection, {}).get('max_abs_flow', 0)
            max_time = regional_max_flows.get(connection, {}).get('time', 'N/A')
            
            utilization = (max_flow / total_capacity * 100) if total_capacity > 0 else 0
            
            print(f"\n{connection}:")
            print(f"  총 용량: {total_capacity:,.1f} MW")
            print(f"  최대 송전: {max_flow:,.1f} MW")
            print(f"  최대 이용률: {utilization:.1f}%")
            print(f"  최대 시점: {max_time}")
            
            if utilization > 90:  # 90% 이상 사용
                capacity_issues.append({
                    'connection': connection,
                    'capacity': total_capacity,
                    'max_flow': max_flow,
                    'utilization': utilization,
                    'time': max_time
                })
                print(f"  ⚠️ 고위험: {utilization:.1f}% 이용률")
            elif utilization > 70:  # 70% 이상 사용
                print(f"  ⚠️ 주의: {utilization:.1f}% 이용률")
            else:
                print(f"  ✅ 정상: {utilization:.1f}% 이용률")
        
        # 6. 가장 문제가 되는 지역 식별
        if capacity_issues:
            print(f"\n=== 송전 병목 위험 지역 TOP 5 ===")
            capacity_issues.sort(key=lambda x: x['utilization'], reverse=True)
            
            for i, issue in enumerate(capacity_issues[:5]):
                print(f"{i+1}. {issue['connection']}")
                print(f"   이용률: {issue['utilization']:.1f}%")
                print(f"   부족량: {issue['max_flow'] - issue['capacity']:,.1f} MW")
                print(f"   시점: {issue['time']}")
                print()
            
            return capacity_issues
        else:
            print("\n✅ 모든 지역간 연결이 충분한 용량을 가지고 있습니다.")
            print("송전 병목이 아닌 다른 원인을 찾아야 합니다:")
            print("- 개별 버스의 발전/수요 불균형")
            print("- 시간별 변동성")
            print("- 링크(CHP, HP 등) 제약")
            print("- 저장장치 운영 제약")
            
            return []
            
    except Exception as e:
        print(f"분석 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return []

if __name__ == "__main__":
    reanalyze_transmission_properly() 