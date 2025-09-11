import pandas as pd
import numpy as np
import shutil
from datetime import datetime

def fix_infeasibility_simple():
    """최적화 실행 불가능성 해결 (간단한 버전)"""
    
    print("=== 최적화 실행 불가능성 해결 (간단한 버전) ===")
    
    # 백업 생성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f'integrated_input_data_backup_simple_{timestamp}.xlsx'
    shutil.copy2('integrated_input_data.xlsx', backup_filename)
    print(f"백업 파일 생성: {backup_filename}")
    
    try:
        # 발전기 데이터만 로드하고 수정
        print("발전기 데이터 로드 중...")
        generators_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators', engine='openpyxl')
        print(f"원본 발전기 수: {len(generators_df)}")
        
        # 1. 부족 지역에 비상 LNG 발전기 추가
        deficit_regions = ['DGU', 'DJN', 'GWJ', 'SEL']
        deficit_amounts = {
            'DGU': 2666,  # 2222 * 1.2
            'DJN': 1966,  # 1639 * 1.2  
            'GWJ': 2060,  # 1717 * 1.2
            'SEL': 6265   # 5221 * 1.2
        }
        
        print(f"부족 지역: {deficit_regions}")
        
        # 새 발전기 추가
        new_generators = []
        for region, capacity in deficit_amounts.items():
            new_gen = {
                'name': f'{region}_Emergency_LNG',
                'bus': f'{region}_{region}_EL',
                'carrier': 'electricity',
                'p_nom': capacity,
                'p_nom_extendable': False,
                'marginal_cost': 100.0,
                'efficiency': 0.45,
                'capital_cost': 0,
                'p_min_pu': 0.0
            }
            new_generators.append(new_gen)
            print(f"{region} 지역에 비상 LNG 발전기 추가: {capacity} MW")
        
        # 2. 기존 발전기의 최소 출력 제약 완화
        print("\n기존 발전기 최소 출력 제약 완화:")
        
        # p_min_pu 컬럼이 없으면 추가
        if 'p_min_pu' not in generators_df.columns:
            generators_df['p_min_pu'] = 0.0
            print("p_min_pu 컬럼 추가됨")
        
        # 모든 발전기의 최소 출력을 0으로 설정
        generators_df['p_min_pu'] = 0.0
        
        # 원자력 발전기의 최소 출력을 30%로 설정
        nuclear_mask = generators_df['name'].str.contains('Nuclear', na=False)
        generators_df.loc[nuclear_mask, 'p_min_pu'] = 0.3
        print(f"원자력 발전기 {nuclear_mask.sum()}개의 최소 출력을 30%로 설정")
        
        # 석탄 발전기의 최소 출력을 20%로 설정
        coal_mask = generators_df['name'].str.contains('Coal', na=False)
        generators_df.loc[coal_mask, 'p_min_pu'] = 0.2
        print(f"석탄 발전기 {coal_mask.sum()}개의 최소 출력을 20%로 설정")
        
        # 3. 새 발전기를 기존 DataFrame에 추가
        new_gen_df = pd.DataFrame(new_generators)
        generators_df = pd.concat([generators_df, new_gen_df], ignore_index=True)
        
        print(f"\n수정 후 발전기 수: {len(generators_df)}")
        print(f"추가된 비상 발전기: {len(new_generators)}개")
        
        # 4. 다른 시트들 로드
        print("\n다른 시트들 로드 중...")
        other_sheets = {}
        sheet_names = ['buses', 'loads', 'stores', 'links', 'lines', 'timeseries', 'renewable_patterns', 'load_patterns']
        
        for sheet_name in sheet_names:
            try:
                other_sheets[sheet_name] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet_name, engine='openpyxl')
                print(f"시트 로드: {sheet_name} ({len(other_sheets[sheet_name])} 행)")
            except Exception as e:
                print(f"시트 {sheet_name} 로드 실패: {e}")
        
        # 5. 수정된 데이터 저장
        print("\n수정된 데이터 저장 중...")
        
        with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
            # 수정된 발전기 데이터 저장
            generators_df.to_excel(writer, sheet_name='generators', index=False)
            print("발전기 시트 저장 완료")
            
            # 다른 시트들 저장
            for sheet_name, df in other_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"시트 저장: {sheet_name}")
        
        print("수정된 데이터 저장 완료!")
        
        # 6. 수정 사항 요약
        print("\n=== 수정 사항 요약 ===")
        print("1. 부족 지역에 비상 LNG 발전기 추가:")
        for region, capacity in deficit_amounts.items():
            print(f"   - {region}: {capacity} MW")
        
        print("2. 발전기 최소 출력 제약 완화:")
        print("   - 모든 발전기: 0%")
        print("   - 원자력: 30%")
        print("   - 석탄: 20%")
        
        total_additional = sum(deficit_amounts.values())
        print(f"3. 총 추가 발전 용량: {total_additional:,} MW")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        # 백업에서 복원
        shutil.copy2(backup_filename, 'integrated_input_data.xlsx')
        print("백업에서 복원됨")
        return False

if __name__ == "__main__":
    success = fix_infeasibility_simple()
    if success:
        print("\n✅ 최적화 실행 불가능성 해결 완료!")
        print("이제 PyPSA_GUI.py를 다시 실행해보세요.")
    else:
        print("\n❌ 문제 해결 중 오류가 발생했습니다.") 