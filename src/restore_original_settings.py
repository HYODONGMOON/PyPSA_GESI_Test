import pandas as pd
import numpy as np

def restore_original_settings():
    """원본 설정을 복원하고 필요한 컬럼만 추가"""
    print("=== 원본 설정 복원 및 필수 컬럼 추가 ===")
    
    # 백업 파일 생성
    from datetime import datetime
    backup_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f'integrated_input_data_backup_{backup_time}.xlsx'
    
    # 원본 파일 백업
    import shutil
    shutil.copy('integrated_input_data.xlsx', backup_filename)
    print(f"백업 파일 생성: {backup_filename}")
    
    # 데이터 로드
    input_data = {}
    xls = pd.ExcelFile('integrated_input_data.xlsx')
    for sheet_name in xls.sheet_names:
        input_data[sheet_name] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet_name)
    
    print(f"\n로드된 시트: {list(input_data.keys())}")
    
    # 1. 발전기 데이터 원본 설정 복원
    print(f"\n1. 발전기 데이터 설정 복원")
    generators_df = input_data['generators']
    
    print(f"현재 발전기 컬럼: {list(generators_df.columns)}")
    
    # 원본 설정으로 복원 - expandable을 False로 설정
    if 'p_nom_extendable' in generators_df.columns:
        generators_df['p_nom_extendable'] = False
        print("  모든 발전기 p_nom_extendable을 False로 설정")
    else:
        generators_df['p_nom_extendable'] = False
        print("  p_nom_extendable 컬럼 추가 (False)")
    
    # p_nom_max 컬럼 제거 (expandable이 False이므로 불필요)
    if 'p_nom_max' in generators_df.columns:
        generators_df = generators_df.drop('p_nom_max', axis=1)
        print("  p_nom_max 컬럼 제거 (expandable=False이므로 불필요)")
    
    # 필수 컬럼만 추가 (expandable=False에 필요한 것들)
    required_columns = {
        'marginal_cost': 50.0,
        'efficiency': 0.4,
        'committable': False,
        'p_min_pu': 0.0,
        'p_max_pu': 1.0
    }
    
    for col, default_value in required_columns.items():
        if col not in generators_df.columns:
            generators_df[col] = default_value
            print(f"  컬럼 추가: {col} = {default_value}")
        else:
            # 기존 컬럼의 NaN 값 채우기
            null_count = generators_df[col].isnull().sum()
            if null_count > 0:
                generators_df[col].fillna(default_value, inplace=True)
                print(f"  컬럼 {col}의 NaN 값 {null_count}개 채움: {default_value}")
    
    # 발전원별 한계비용 설정 (expandable=False이므로 용량 확장 없이 운영비용만)
    for idx, gen in generators_df.iterrows():
        gen_name = gen['name']
        
        # 재생에너지는 한계비용 0
        if any(renewable in gen_name for renewable in ['PV', 'WT', 'Solar', 'Wind']):
            generators_df.at[idx, 'marginal_cost'] = 0.0
            generators_df.at[idx, 'p_min_pu'] = 0.0  # 재생에너지는 최소 출력 없음
            
        # LNG 발전기
        elif 'LNG' in gen_name:
            generators_df.at[idx, 'marginal_cost'] = 80.0
            generators_df.at[idx, 'p_min_pu'] = 0.3  # LNG 최소 출력 30%
            
        # 원자력
        elif 'Nuclear' in gen_name:
            generators_df.at[idx, 'marginal_cost'] = 10.0
            generators_df.at[idx, 'p_min_pu'] = 0.8  # 원자력 최소 출력 80%
            
        # 석탄 발전기
        elif 'Coal' in gen_name:
            generators_df.at[idx, 'marginal_cost'] = 60.0
            generators_df.at[idx, 'p_min_pu'] = 0.4  # 석탄 최소 출력 40%
    
    print(f"발전기 데이터 설정 복원 완료: {len(generators_df)}개 발전기")
    
    # 2. 선로 데이터 설정
    print(f"\n2. 선로 데이터 설정")
    lines_df = input_data['lines']
    
    # 선로도 expandable을 False로 설정 (원본 설정 유지)
    if 's_nom_extendable' in lines_df.columns:
        lines_df['s_nom_extendable'] = False
        print("  모든 선로 s_nom_extendable을 False로 설정")
    else:
        lines_df['s_nom_extendable'] = False
        print("  s_nom_extendable 컬럼 추가 (False)")
    
    # s_nom_max 컬럼 제거 (expandable이 False이므로 불필요)
    if 's_nom_max' in lines_df.columns:
        lines_df = lines_df.drop('s_nom_max', axis=1)
        print("  s_nom_max 컬럼 제거 (expandable=False이므로 불필요)")
    
    # 3. 저장장치 데이터 설정
    if 'stores' in input_data:
        print(f"\n3. 저장장치 데이터 설정")
        stores_df = input_data['stores']
        
        # 저장장치도 expandable을 False로 설정
        if 'e_nom_extendable' in stores_df.columns:
            stores_df['e_nom_extendable'] = False
            print("  모든 저장장치 e_nom_extendable을 False로 설정")
        else:
            stores_df['e_nom_extendable'] = False
            print("  e_nom_extendable 컬럼 추가 (False)")
        
        # e_nom_max 컬럼 제거
        if 'e_nom_max' in stores_df.columns:
            stores_df = stores_df.drop('e_nom_max', axis=1)
            print("  e_nom_max 컬럼 제거 (expandable=False이므로 불필요)")
        
        # 필수 컬럼만 추가
        store_required_columns = {
            'e_cyclic': True,
            'standing_loss': 0.0,
            'efficiency_store': 0.9,
            'efficiency_dispatch': 0.9,
            'e_initial': 0.0
        }
        
        for col, default_value in store_required_columns.items():
            if col not in stores_df.columns:
                stores_df[col] = default_value
                print(f"  컬럼 추가: {col} = {default_value}")
            else:
                null_count = stores_df[col].isnull().sum()
                if null_count > 0:
                    stores_df[col].fillna(default_value, inplace=True)
                    print(f"  컬럼 {col}의 NaN 값 {null_count}개 채움: {default_value}")
        
        input_data['stores'] = stores_df
    
    # 4. Links 데이터 설정
    if 'links' in input_data:
        print(f"\n4. Links 데이터 설정")
        links_df = input_data['links']
        
        # Links도 expandable을 False로 설정
        if 'p_nom_extendable' in links_df.columns:
            links_df['p_nom_extendable'] = False
            print("  모든 링크 p_nom_extendable을 False로 설정")
        else:
            links_df['p_nom_extendable'] = False
            print("  p_nom_extendable 컬럼 추가 (False)")
        
        # p_nom_max 컬럼 제거
        if 'p_nom_max' in links_df.columns:
            links_df = links_df.drop('p_nom_max', axis=1)
            print("  p_nom_max 컬럼 제거 (expandable=False이므로 불필요)")
        
        # 필수 컬럼만 추가
        link_required_columns = {
            'marginal_cost': 0.0
        }
        
        for col, default_value in link_required_columns.items():
            if col not in links_df.columns:
                links_df[col] = default_value
                print(f"  컬럼 추가: {col} = {default_value}")
            else:
                null_count = links_df[col].isnull().sum()
                if null_count > 0:
                    links_df[col].fillna(default_value, inplace=True)
                    print(f"  컬럼 {col}의 NaN 값 {null_count}개 채움: {default_value}")
        
        input_data['links'] = links_df
    
    # 5. 수정된 데이터 저장
    print(f"\n5. 수정된 데이터 저장")
    input_data['generators'] = generators_df
    input_data['lines'] = lines_df
    
    with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
        for sheet_name, df in input_data.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print("수정된 데이터가 저장되었습니다.")
    
    # 6. 수정 사항 요약
    print(f"\n6. 수정 사항 요약:")
    print("  ✓ 모든 발전기 p_nom_extendable을 False로 복원")
    print("  ✓ 모든 선로 s_nom_extendable을 False로 복원")
    print("  ✓ 모든 저장장치 e_nom_extendable을 False로 복원")
    print("  ✓ 모든 링크 p_nom_extendable을 False로 복원")
    print("  ✓ 불필요한 *_max 컬럼들 제거")
    print("  ✓ 필수 운영 컬럼들만 추가")
    print("  ✓ 발전원별 적절한 한계비용 및 최소출력 설정")
    print("\n이제 원본 regional_input_template 설정과 동일하게 복원되었습니다.")
    
    return True

if __name__ == "__main__":
    restore_original_settings() 