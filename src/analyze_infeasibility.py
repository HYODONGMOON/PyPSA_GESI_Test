import pandas as pd
import numpy as np
import pypsa

def analyze_infeasibility():
    """최적화 불가능 문제 분석"""
    print("=== 최적화 불가능 문제 분석 ===\n")
    
    try:
        # 데이터 로드
        input_data = {}
        xls = pd.ExcelFile('integrated_input_data.xlsx')
        
        for sheet_name in xls.sheet_names:
            input_data[sheet_name] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet_name)
        
        print("1. 데이터 기본 정보:")
        for sheet_name, df in input_data.items():
            print(f"   - {sheet_name}: {len(df)}행")
        
        # 발전-부하 균형 분석
        print("\n2. 발전-부하 균형 분석:")
        
        # 총 발전 용량
        total_gen_capacity = input_data['generators']['p_nom'].sum()
        print(f"   총 발전 용량: {total_gen_capacity:,.0f} MW")
        
        # 총 부하
        total_load = input_data['loads']['p_set'].sum()
        print(f"   총 부하: {total_load:,.0f} MW")
        
        balance_ratio = total_gen_capacity / total_load
        print(f"   발전/부하 비율: {balance_ratio:.2f}")
        
        if balance_ratio < 1.0:
            print("   ⚠️ 경고: 발전 용량이 부하보다 부족합니다!")
        
        # 지역별 발전-부하 균형
        print("\n3. 지역별 발전-부하 균형:")
        
        # 지역별 발전 용량
        regional_gen = {}
        for _, gen in input_data['generators'].iterrows():
            if '_' in gen['bus']:
                region = gen['bus'].split('_')[0]
                if region not in regional_gen:
                    regional_gen[region] = 0
                regional_gen[region] += gen['p_nom']
        
        # 지역별 부하
        regional_load = {}
        for _, load in input_data['loads'].iterrows():
            if '_' in load['bus']:
                region = load['bus'].split('_')[0]
                if region not in regional_load:
                    regional_load[region] = 0
                regional_load[region] += load['p_set']
        
        print("   지역별 발전/부하 비율:")
        imbalanced_regions = []
        for region in sorted(set(list(regional_gen.keys()) + list(regional_load.keys()))):
            gen_cap = regional_gen.get(region, 0)
            load_demand = regional_load.get(region, 0)
            if load_demand > 0:
                ratio = gen_cap / load_demand
                print(f"   - {region}: {ratio:.2f} (발전: {gen_cap:,.0f} MW, 부하: {load_demand:,.0f} MW)")
                if ratio < 0.5:  # 발전 용량이 부하의 50% 미만
                    imbalanced_regions.append(region)
            else:
                print(f"   - {region}: 무한대 (발전: {gen_cap:,.0f} MW, 부하: 0 MW)")
        
        if imbalanced_regions:
            print(f"   ⚠️ 발전 부족 지역: {imbalanced_regions}")
        
        # 네트워크 연결성 분석
        print("\n4. 네트워크 연결성 분석:")
        
        # 선로가 있는지 확인
        if 'lines' in input_data and not input_data['lines'].empty:
            lines_df = input_data['lines']
            print(f"   선로 수: {len(lines_df)}")
            
            # 연결된 지역 확인
            connected_regions = set()
            for _, line in lines_df.iterrows():
                bus0_region = line['bus0'].split('_')[0] if '_' in str(line['bus0']) else str(line['bus0'])
                bus1_region = line['bus1'].split('_')[0] if '_' in str(line['bus1']) else str(line['bus1'])
                connected_regions.add(bus0_region)
                connected_regions.add(bus1_region)
            
            print(f"   연결된 지역: {sorted(connected_regions)}")
            
            # 고립된 지역 확인
            all_regions = set()
            for _, bus in input_data['buses'].iterrows():
                if '_' in bus['name']:
                    region = bus['name'].split('_')[0]
                    all_regions.add(region)
            
            isolated_regions = all_regions - connected_regions
            if isolated_regions:
                print(f"   ⚠️ 고립된 지역: {sorted(isolated_regions)}")
            else:
                print("   모든 지역이 연결되어 있습니다.")
        else:
            print("   ⚠️ 선로가 없습니다! 모든 지역이 고립되어 있습니다.")
        
        # 재생에너지 패턴 분석
        print("\n5. 재생에너지 패턴 분석:")
        if 'renewable_patterns' in input_data:
            patterns_df = input_data['renewable_patterns']
            print(f"   패턴 데이터 행 수: {len(patterns_df)}")
            print(f"   패턴 컬럼: {list(patterns_df.columns)}")
            
            # 패턴 값 범위 확인
            for col in patterns_df.columns:
                if patterns_df[col].dtype in ['float64', 'int64']:
                    min_val = patterns_df[col].min()
                    max_val = patterns_df[col].max()
                    mean_val = patterns_df[col].mean()
                    print(f"   - {col}: 범위 [{min_val:.3f}, {max_val:.3f}], 평균 {mean_val:.3f}")
                    
                    if max_val == 0:
                        print(f"     ⚠️ {col} 패턴이 모두 0입니다!")
        
        # 저장장치 분석
        print("\n6. 저장장치 분석:")
        if 'stores' in input_data:
            stores_df = input_data['stores']
            print(f"   저장장치 수: {len(stores_df)}")
            
            # 저장장치 타입별 분석
            store_types = stores_df['carrier'].value_counts()
            print("   타입별 저장장치:")
            for store_type, count in store_types.items():
                print(f"   - {store_type}: {count}개")
            
            # 저장 용량 분석
            total_storage = stores_df['e_nom'].sum()
            print(f"   총 저장 용량: {total_storage:,.0f} MWh")
            
            # 확장 가능한 저장장치
            extendable_stores = stores_df[stores_df['e_nom_extendable'] == True]
            print(f"   확장 가능한 저장장치: {len(extendable_stores)}개")
        
        # Links 분석
        print("\n7. Links 분석:")
        if 'links' in input_data:
            links_df = input_data['links']
            print(f"   Link 수: {len(links_df)}")
            
            # 효율성 분석
            avg_efficiency = links_df['efficiency'].mean()
            print(f"   평균 효율: {avg_efficiency:.2f}")
            
            # 용량 분석
            total_link_capacity = links_df['p_nom'].sum()
            print(f"   총 Link 용량: {total_link_capacity:,.0f} MW")
        
        # 제약조건 분석
        print("\n8. 제약조건 분석:")
        if 'constraints' in input_data:
            constraints_df = input_data['constraints']
            print(f"   제약조건 수: {len(constraints_df)}")
            for _, constraint in constraints_df.iterrows():
                print(f"   - {constraint['name']}: {constraint['type']} {constraint['sense']} {constraint['constant']}")
        
        # 권장사항
        print("\n9. 권장사항:")
        recommendations = []
        
        if balance_ratio < 1.0:
            recommendations.append("발전 용량을 늘리거나 부하를 줄이세요.")
        
        if imbalanced_regions:
            recommendations.append(f"발전 부족 지역({imbalanced_regions})에 발전기를 추가하거나 송전선을 강화하세요.")
        
        if 'lines' not in input_data or input_data['lines'].empty:
            recommendations.append("지역 간 송전선을 추가하세요.")
        
        if isolated_regions:
            recommendations.append(f"고립된 지역({sorted(isolated_regions)})을 송전망에 연결하세요.")
        
        if not recommendations:
            recommendations.append("데이터 구조상 문제가 없어 보입니다. 최적화 파라미터나 솔버 설정을 확인해보세요.")
        
        for i, rec in enumerate(recommendations, 1):
            print(f"   {i}. {rec}")
        
        return True
        
    except Exception as e:
        print(f"분석 중 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    analyze_infeasibility() 