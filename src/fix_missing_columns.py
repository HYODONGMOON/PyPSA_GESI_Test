import pandas as pd
import numpy as np

def fix_missing_columns():
    """누락된 필수 컬럼들을 추가하여 최적화 문제 해결"""
    print("=== 누락된 필수 컬럼 추가 ===")
    
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
    
    # 1. 발전기 데이터에 누락된 컬럼 추가
    print(f"\n1. 발전기 데이터 컬럼 추가")
    generators_df = input_data['generators']
    
    # 필수 컬럼들 확인 및 추가
    required_columns = {
        'p_nom_extendable': True,
        'p_nom_min': 0.0,
        'p_nom_max': 50000.0,  # 충분히 큰 값으로 설정
        'marginal_cost': 50.0,
        'capital_cost': 1000000.0,
        'efficiency': 0.4,
        'committable': False,
        'ramp_limit_up': np.nan,
        'ramp_limit_down': np.nan,
        'min_up_time': 0,
        'min_down_time': 0,
        'start_up_cost': 0.0,
        'shut_down_cost': 0.0
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
    
    # 발전기별 특별 설정
    for idx, gen in generators_df.iterrows():
        gen_name = gen['name']
        
        # 재생에너지는 확장 가능하게 설정
        if any(renewable in gen_name for renewable in ['PV', 'WT', 'Solar', 'Wind']):
            generators_df.at[idx, 'p_nom_extendable'] = True
            generators_df.at[idx, 'p_nom_max'] = generators_df.at[idx, 'p_nom'] * 10  # 10배까지 확장
            generators_df.at[idx, 'marginal_cost'] = 0.0  # 재생에너지는 한계비용 0
            
        # LNG 발전기는 확장 가능하게 설정
        elif 'LNG' in gen_name:
            generators_df.at[idx, 'p_nom_extendable'] = True
            generators_df.at[idx, 'p_nom_max'] = generators_df.at[idx, 'p_nom'] * 5  # 5배까지 확장
            generators_df.at[idx, 'marginal_cost'] = 80.0  # LNG 한계비용
            
        # 원자력은 고정 용량
        elif 'Nuclear' in gen_name:
            generators_df.at[idx, 'p_nom_extendable'] = False
            generators_df.at[idx, 'marginal_cost'] = 10.0  # 원자력 한계비용
            
        # 석탄 발전기
        elif 'Coal' in gen_name:
            generators_df.at[idx, 'p_nom_extendable'] = True
            generators_df.at[idx, 'p_nom_max'] = generators_df.at[idx, 'p_nom'] * 2  # 2배까지 확장
            generators_df.at[idx, 'marginal_cost'] = 60.0  # 석탄 한계비용
    
    print(f"발전기 데이터 컬럼 추가 완료: {len(generators_df)}개 발전기")
    
    # 2. 선로 데이터에 누락된 컬럼 추가
    print(f"\n2. 선로 데이터 컬럼 추가")
    lines_df = input_data['lines']
    
    line_required_columns = {
        's_nom_extendable': True,
        's_nom_min': 0.0,
        's_nom_max': 100000.0,  # 충분히 큰 값
        'capital_cost': 1000000.0,
        'length': 100.0,
        'terrain_factor': 1.0
    }
    
    for col, default_value in line_required_columns.items():
        if col not in lines_df.columns:
            lines_df[col] = default_value
            print(f"  컬럼 추가: {col} = {default_value}")
        else:
            null_count = lines_df[col].isnull().sum()
            if null_count > 0:
                lines_df[col].fillna(default_value, inplace=True)
                print(f"  컬럼 {col}의 NaN 값 {null_count}개 채움: {default_value}")
    
    # 3. 저장장치 데이터에 누락된 컬럼 추가
    if 'stores' in input_data:
        print(f"\n3. 저장장치 데이터 컬럼 추가")
        stores_df = input_data['stores']
        
        store_required_columns = {
            'e_nom_extendable': True,
            'e_nom_min': 0.0,
            'e_nom_max': 100000.0,
            'e_cyclic': True,
            'standing_loss': 0.0,
            'efficiency_store': 0.9,
            'efficiency_dispatch': 0.9,
            'e_initial': 0.0,
            'capital_cost': 500000.0
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
    
    # 4. Links 데이터에 누락된 컬럼 추가
    if 'links' in input_data:
        print(f"\n4. Links 데이터 컬럼 추가")
        links_df = input_data['links']
        
        link_required_columns = {
            'p_nom_extendable': True,
            'p_nom_min': 0.0,
            'p_nom_max': 100000.0,
            'capital_cost': 1000000.0,
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
    print("  ✓ 발전기 데이터에 필수 컬럼 추가 (p_nom_max, p_nom_extendable 등)")
    print("  ✓ 선로 데이터에 필수 컬럼 추가 (s_nom_max, s_nom_extendable 등)")
    print("  ✓ 저장장치 데이터에 필수 컬럼 추가")
    print("  ✓ Links 데이터에 필수 컬럼 추가")
    print("  ✓ 재생에너지 발전기 확장성 강화")
    print("  ✓ LNG 발전기 확장성 강화")
    print("  ✓ 모든 NaN 값 적절한 기본값으로 채움")
    
    return True

if __name__ == "__main__":
    fix_missing_columns() 