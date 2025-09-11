import pandas as pd
import numpy as np
import shutil
from datetime import datetime

def fix_infeasibility():
    """최적화 실행 불가능성 해결"""
    
    print("=== 최적화 실행 불가능성 해결 ===")
    
    # 백업 생성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f'integrated_input_data_backup_infeasibility_{timestamp}.xlsx'
    shutil.copy2('integrated_input_data.xlsx', backup_filename)
    print(f"백업 파일 생성: {backup_filename}")
    
    # 데이터 로드
    generators_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
    loads_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='loads')
    
    print(f"원본 발전기 수: {len(generators_df)}")
    
    # 1. 부족 지역 확인
    deficit_regions = ['DGU', 'DJN', 'GWJ', 'SEL']
    deficit_amounts = {
        'DGU': 2222,
        'DJN': 1639, 
        'GWJ': 1717,
        'SEL': 5221
    }
    
    print(f"부족 지역: {deficit_regions}")
    
    # 2. 부족 지역에 임시 LNG 발전기 추가
    new_generators = []
    
    for region in deficit_regions:
        deficit_amount = deficit_amounts[region]
        # 부족량의 120%로 여유있게 설정
        additional_capacity = int(deficit_amount * 1.2)
        
        new_gen = {
            'name': f'{region}_Emergency_LNG',
            'bus': f'{region}_{region}_EL',
            'carrier': 'electricity',
            'p_nom': additional_capacity,
            'p_nom_extendable': False,
            'marginal_cost': 100.0,  # 높은 비용으로 설정 (마지막 수단)
            'p_min_pu': 0.0,  # 최소 출력 0%
            'efficiency': 0.45,
            'capital_cost': 0
        }
        
        new_generators.append(new_gen)
        print(f"{region} 지역에 비상 LNG 발전기 추가: {additional_capacity} MW")
    
    # 3. 기존 발전기의 최소 출력 제약 완화
    print("\n기존 발전기 최소 출력 제약 완화:")
    
    # p_min_pu 컬럼이 없으면 추가
    if 'p_min_pu' not in generators_df.columns:
        generators_df['p_min_pu'] = 0.0
        print("p_min_pu 컬럼 추가됨")
    
    # 모든 발전기의 최소 출력을 0으로 설정
    generators_df['p_min_pu'] = 0.0
    
    # 원자력 발전기의 최소 출력을 낮춤 (원래 70% -> 30%)
    nuclear_mask = generators_df['name'].str.contains('Nuclear', na=False)
    generators_df.loc[nuclear_mask, 'p_min_pu'] = 0.3
    print(f"원자력 발전기 {nuclear_mask.sum()}개의 최소 출력을 30%로 설정")
    
    # 석탄 발전기의 최소 출력을 낮춤 (원래 40% -> 20%)
    coal_mask = generators_df['name'].str.contains('Coal', na=False)
    generators_df.loc[coal_mask, 'p_min_pu'] = 0.2
    print(f"석탄 발전기 {coal_mask.sum()}개의 최소 출력을 20%로 설정")
    
    # 4. 새 발전기를 기존 DataFrame에 추가
    new_gen_df = pd.DataFrame(new_generators)
    generators_df = pd.concat([generators_df, new_gen_df], ignore_index=True)
    
    print(f"\n수정 후 발전기 수: {len(generators_df)}")
    print(f"추가된 비상 발전기: {len(new_generators)}개")
    
    # 5. 부하 패턴 단순화 (일정한 부하로 변경)
    print("\n부하 패턴 단순화:")
    
    # 전력 부하만 처리
    el_loads = loads_df[loads_df['name'].str.contains('_Demand_EL', na=False)]
    print(f"전력 부하 수: {len(el_loads)}")
    
    # 6. 데이터 저장
    with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
        # 기존 시트들 다시 로드해서 저장
        original_file = pd.ExcelFile('integrated_input_data.xlsx')
        
        for sheet_name in original_file.sheet_names:
            if sheet_name == 'generators':
                generators_df.to_excel(writer, sheet_name='generators', index=False)
            else:
                df = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet_name)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print("\n수정된 데이터 저장 완료")
    
    # 7. 수정 사항 요약
    print("\n=== 수정 사항 요약 ===")
    print("1. 부족 지역에 비상 LNG 발전기 추가:")
    for region, amount in deficit_amounts.items():
        additional = int(amount * 1.2)
        print(f"   - {region}: {additional} MW")
    
    print("2. 발전기 최소 출력 제약 완화:")
    print("   - 모든 발전기: 0%")
    print("   - 원자력: 30%")
    print("   - 석탄: 20%")
    
    print("3. 총 추가 발전 용량:", sum(int(v * 1.2) for v in deficit_amounts.values()), "MW")
    
    return True

if __name__ == "__main__":
    fix_infeasibility() 