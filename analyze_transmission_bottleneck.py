import pandas as pd
import numpy as np

def analyze_transmission_bottleneck():
    print("=== 송전망 병목 현상 분석 ===")
    
    try:
        # 데이터 로드
        buses_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='buses')
        loads_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='loads')
        generators_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
        lines_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='lines')
        
        print(f"버스: {len(buses_df)}개, 부하: {len(loads_df)}개, 발전기: {len(generators_df)}개, 송전선로: {len(lines_df)}개")
        
        # 전력 버스만 필터링 (carrier가 AC인 것들)
        power_buses = buses_df[buses_df['carrier'] == 'AC']['name'].tolist()
        print(f"\n전력 버스: {len(power_buses)}개")
        print(power_buses)
        
        # 지역별 분석
        print("\n=== 지역별 전력 수급 분석 ===")
        
        regional_analysis = {}
        
        for bus in power_buses:
            region = bus.replace('_EL', '') if '_EL' in bus else bus
            
            # 해당 지역의 전력 수요 (loads)
            region_loads = loads_df[loads_df['bus'] == bus]
            total_demand = region_loads['p_set'].sum() if not region_loads.empty else 0
            
            # 해당 지역의 발전 용량 (generators)
            region_gens = generators_df[generators_df['bus'] == bus]
            # 슬랙 발전기 제외
            region_gens_no_slack = region_gens[~region_gens['name'].str.contains('Slack|Power_Slack', na=False)]
            total_generation = region_gens_no_slack['p_nom'].sum() if not region_gens_no_slack.empty else 0
            
            # 해당 지역으로 들어오는 송전선로 용량
            incoming_lines = lines_df[lines_df['bus1'] == bus]
            outgoing_lines = lines_df[lines_df['bus0'] == bus]
            
            total_incoming_capacity = incoming_lines['s_nom'].sum() if not incoming_lines.empty else 0
            total_outgoing_capacity = outgoing_lines['s_nom'].sum() if not outgoing_lines.empty else 0
            total_transmission_capacity = total_incoming_capacity + total_outgoing_capacity
            
            # 전력 부족량
            deficit = total_demand - total_generation
            
            # 가능한 최대 수급 (자체 생산 + 송전 가능)
            max_available = total_generation + total_transmission_capacity
            
            # 문제 여부 판단
            is_problematic = deficit > total_transmission_capacity if deficit > 0 else False
            
            regional_analysis[region] = {
                'demand': total_demand,
                'generation': total_generation,
                'deficit': deficit,
                'transmission_capacity': total_transmission_capacity,
                'max_available': max_available,
                'is_problematic': is_problematic,
                'incoming_capacity': total_incoming_capacity,
                'outgoing_capacity': total_outgoing_capacity
            }
            
            status = "⚠️ 문제" if is_problematic else "✅ 정상"
            print(f"\n{region} ({status}):")
            print(f"  전력 수요: {total_demand:,.1f} MW")
            print(f"  자체 발전: {total_generation:,.1f} MW")
            print(f"  부족량: {deficit:,.1f} MW")
            print(f"  송전 용량: {total_transmission_capacity:,.1f} MW (들어오는: {total_incoming_capacity:,.1f}, 나가는: {total_outgoing_capacity:,.1f})")
            print(f"  최대 가능: {max_available:,.1f} MW")
            
            if is_problematic:
                shortage = deficit - total_transmission_capacity
                print(f"  🚨 송전 부족: {shortage:,.1f} MW (infeasible 원인 가능)")
        
        # 문제 지역 요약
        problematic_regions = [region for region, data in regional_analysis.items() if data['is_problematic']]
        
        print(f"\n=== 요약 ===")
        print(f"문제 지역: {len(problematic_regions)}개")
        if problematic_regions:
            print("문제 지역 목록:")
            for region in problematic_regions:
                data = regional_analysis[region]
                shortage = data['deficit'] - data['transmission_capacity']
                print(f"- {region}: {shortage:,.1f} MW 부족")
        else:
            print("모든 지역이 송전 용량 측면에서 문제없음")
            print("다른 원인을 찾아야 함 (링크 제약, 시간별 변동 등)")
        
        # 가장 취약한 지역 상위 5개
        print(f"\n=== 송전 의존도가 높은 지역 TOP 5 ===")
        vulnerable_regions = []
        for region, data in regional_analysis.items():
            if data['deficit'] > 0:
                dependency_ratio = data['deficit'] / data['demand'] * 100
                vulnerable_regions.append((region, dependency_ratio, data['deficit']))
        
        vulnerable_regions.sort(key=lambda x: x[1], reverse=True)
        for i, (region, ratio, deficit) in enumerate(vulnerable_regions[:5]):
            print(f"{i+1}. {region}: {ratio:.1f}% 송전 의존 ({deficit:,.1f} MW 부족)")
        
        return problematic_regions
        
    except Exception as e:
        print(f"분석 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return []

if __name__ == "__main__":
    analyze_transmission_bottleneck() 