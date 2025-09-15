import pandas as pd
import numpy as np

def analyze_infeasible_problem():
    """infeasible 문제를 분석하는 함수"""
    
    print("=== Infeasible 문제 분석 ===")
    
    try:
        # 파일 로드
        xls = pd.ExcelFile('integrated_input_data.xlsx')
        
        # 기본 데이터 로드
        constraints = pd.read_excel('integrated_input_data.xlsx', sheet_name='constraints')
        loads = pd.read_excel('integrated_input_data.xlsx', sheet_name='loads')
        gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
        links = pd.read_excel('integrated_input_data.xlsx', sheet_name='links')
        
        print("\n1. 제약조건 분석:")
        print(constraints)
        
        # CO2 제약 찾기
        if 'name' in constraints.columns:
            co2_constraints = constraints[constraints['name'].astype(str).str.contains('CO2', na=False)]
            if not co2_constraints.empty:
                print("\nCO2 제약조건:")
                for _, row in co2_constraints.iterrows():
                    print(f"- {row['name']}: {row.get('constant', 'N/A')} (조건: {row.get('sense', 'N/A')})")
        
        print("\n2. 수급 균형 분석:")
        total_load = loads['p_set'].sum()
        total_gen = gens['p_nom'].sum()
        
        print(f"총 부하: {total_load:,.0f} MW")
        print(f"총 발전용량: {total_gen:,.0f} MW")
        print(f"발전 여유율: {(total_gen / total_load - 1) * 100:.1f}%")
        
        if total_gen < total_load:
            print("⚠️ 경고: 총 발전용량이 총 부하보다 부족!")
        
        print("\n3. CHP 링크 분석:")
        chp_links = links[links['name'].str.contains('CHP', na=False)]
        print(f"CHP 링크 개수: {len(chp_links)}")
        
        if not chp_links.empty:
            print("CHP 링크 예시:")
            for _, link in chp_links.head(5).iterrows():
                eff = link.get('efficiency', 'N/A')
                eff2 = link.get('efficiency2', 'N/A')
                print(f"- {link['name']}: {link['bus0']} -> {link['bus1']} (eff:{eff}, eff2:{eff2})")
        
        print("\n4. 슬랙 발전기 확인:")
        slack_gens = gens[gens['name'].str.contains('Slack|Fallback', na=False)]
        print(f"슬랙/Fallback 발전기 개수: {len(slack_gens)}")
        
        if not slack_gens.empty:
            print("슬랙 발전기 예시:")
            for _, gen in slack_gens.head(5).iterrows():
                extendable = gen.get('p_nom_extendable', False)
                mcost = gen.get('marginal_cost', 'N/A')
                print(f"- {gen['name']}: {gen['p_nom']} MW (확장가능:{extendable}, 비용:{mcost})")
        
        print("\n5. 지역별 수급 균형:")
        # 지역별 분석
        regions = set()
        for name in loads['name']:
            region = str(name).split('_')[0]
            regions.add(region)
        
        for region in sorted(regions)[:10]:  # 상위 10개 지역만
            region_loads = loads[loads['name'].str.startswith(f"{region}_")]
            region_gens = gens[gens['name'].str.startswith(f"{region}_")]
            
            if not region_loads.empty or not region_gens.empty:
                load_sum = region_loads['p_set'].sum() if not region_loads.empty else 0
                gen_sum = region_gens['p_nom'].sum() if not region_gens.empty else 0
                
                balance = "균형" if gen_sum >= load_sum else "부족"
                print(f"{region}: 발전 {gen_sum:.0f} MW, 부하 {load_sum:.0f} MW ({balance})")
                
    except Exception as e:
        print(f"분석 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    analyze_infeasible_problem() 