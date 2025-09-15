import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def analyze_transmission_flow():
    print("=== 송전선로 시간별 전력 흐름 분석 ===")
    
    try:
        # 1. 성공한 케이스의 시간별 전력 흐름 로드
        print("1. 시간별 전력 흐름 데이터 로드 중...")
        line_usage_df = pd.read_csv('results/20250915_151959/optimization_result_20250915_151959_line_usage.csv', index_col=0)
        print(f"   - 시간 단계: {len(line_usage_df)}개")
        print(f"   - 송전선로: {len(line_usage_df.columns)}개")
        
        # 2. 현재 실제 송전선로 용량 로드
        print("\n2. 실제 송전선로 용량 데이터 로드 중...")
        lines_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='lines')
        print(f"   - 송전선로 정보: {len(lines_df)}개")
        
        # 송전선로 용량 딕셔너리 생성
        line_capacity = {}
        for _, line in lines_df.iterrows():
            line_name = line['name']
            capacity = line['s_nom']  # 실제 용량
            line_capacity[line_name] = capacity
        
        print(f"   - 송전선로 용량 정보: {len(line_capacity)}개")
        
        # 3. 시간별 용량 초과 분석
        print("\n3. 시간별 용량 초과 분석 중...")
        
        # 용량 초과 발생 횟수 집계
        capacity_violations = {}
        max_violations = {}
        critical_times = []
        
        for line_name in line_usage_df.columns:
            if line_name in line_capacity:
                actual_capacity = line_capacity[line_name]
                line_flows = line_usage_df[line_name].abs()  # 절댓값으로 방향 무관하게
                
                # 용량 초과 발생 횟수
                violations = (line_flows > actual_capacity).sum()
                if violations > 0:
                    capacity_violations[line_name] = violations
                    max_violations[line_name] = line_flows.max()
                    
                    # 가장 심각한 초과 시점들 기록
                    violation_times = line_flows[line_flows > actual_capacity].index
                    for time_idx in violation_times:
                        critical_times.append({
                            'time': time_idx,
                            'line': line_name,
                            'flow': line_flows[time_idx],
                            'capacity': actual_capacity,
                            'excess': line_flows[time_idx] - actual_capacity,
                            'excess_ratio': (line_flows[time_idx] - actual_capacity) / actual_capacity * 100
                        })
        
        # 4. 결과 분석
        print(f"\n=== 분석 결과 ===")
        print(f"용량 초과 발생 송전선로: {len(capacity_violations)}개")
        
        if capacity_violations:
            print(f"\n📊 용량 초과 송전선로 TOP 10:")
            sorted_violations = sorted(capacity_violations.items(), key=lambda x: x[1], reverse=True)
            
            for i, (line_name, violation_count) in enumerate(sorted_violations[:10]):
                actual_cap = line_capacity[line_name]
                max_flow = max_violations[line_name]
                excess_ratio = (max_flow - actual_cap) / actual_cap * 100
                
                print(f"{i+1:2d}. {line_name}")
                print(f"    용량: {actual_cap:,.0f} MW")
                print(f"    최대 흐름: {max_flow:,.0f} MW")
                print(f"    초과율: {excess_ratio:,.1f}%")
                print(f"    초과 발생: {violation_count}회 ({violation_count/len(line_usage_df)*100:.1f}%)")
                print()
            
            # 5. 가장 심각한 시점들 분석
            print(f"\n⚠️ 가장 심각한 용량 초과 시점 TOP 20:")
            critical_times_sorted = sorted(critical_times, key=lambda x: x['excess_ratio'], reverse=True)
            
            for i, event in enumerate(critical_times_sorted[:20]):
                print(f"{i+1:2d}. 시간: {event['time']}")
                print(f"    송전선로: {event['line']}")
                print(f"    실제 용량: {event['capacity']:,.0f} MW")
                print(f"    필요 흐름: {event['flow']:,.0f} MW")
                print(f"    초과량: {event['excess']:,.0f} MW ({event['excess_ratio']:,.1f}%)")
                print()
            
            # 6. 지역별 문제 분석
            print(f"\n🌏 지역별 송전 병목 분석:")
            regional_issues = {}
            
            for event in critical_times:
                line_name = event['line']
                # 송전선로 이름에서 지역 추출 (예: 'SEL_GGD' -> SEL, GGD)
                if '_' in line_name:
                    regions = line_name.split('_')
                    for region in regions:
                        if region not in regional_issues:
                            regional_issues[region] = {
                                'violation_count': 0,
                                'max_excess': 0,
                                'critical_lines': set()
                            }
                        regional_issues[region]['violation_count'] += 1
                        regional_issues[region]['max_excess'] = max(
                            regional_issues[region]['max_excess'], 
                            event['excess']
                        )
                        regional_issues[region]['critical_lines'].add(line_name)
            
            # 지역별 문제 순위
            regional_sorted = sorted(regional_issues.items(), key=lambda x: x[1]['violation_count'], reverse=True)
            
            for i, (region, issues) in enumerate(regional_sorted[:10]):
                print(f"{i+1:2d}. {region} 지역")
                print(f"    용량 초과 발생: {issues['violation_count']}회")
                print(f"    최대 초과량: {issues['max_excess']:,.0f} MW")
                print(f"    문제 송전선로: {len(issues['critical_lines'])}개")
                print(f"    관련 송전선로: {', '.join(list(issues['critical_lines'])[:3])}...")
                print()
            
            # 7. 시간대별 패턴 분석
            print(f"\n⏰ 시간대별 용량 초과 패턴:")
            hourly_violations = {}
            
            for event in critical_times:
                time_str = str(event['time'])
                # 시간 추출 (예: '2024-01-01 12:00:00' -> 12)
                if ' ' in time_str and ':' in time_str:
                    hour = int(time_str.split(' ')[1].split(':')[0])
                else:
                    hour = 0  # 기본값
                if hour not in hourly_violations:
                    hourly_violations[hour] = 0
                hourly_violations[hour] += 1
            
            if hourly_violations:
                sorted_hours = sorted(hourly_violations.items(), key=lambda x: x[1], reverse=True)
                print("가장 문제가 많은 시간대:")
                for hour, count in sorted_hours[:10]:
                    print(f"  {hour:2d}시: {count}회 용량 초과")
            
            return True
            
        else:
            print("✅ 모든 송전선로가 용량 내에서 운영됨")
            print("   실제 송전선로 용량으로도 문제없이 운영 가능할 것으로 예상")
            print("   다른 원인을 찾아야 함")
            return False
            
    except Exception as e:
        print(f"분석 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    analyze_transmission_flow() 