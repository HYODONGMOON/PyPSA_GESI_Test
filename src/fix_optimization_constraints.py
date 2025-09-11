import pandas as pd
import numpy as np

def fix_optimization_constraints():
    """최적화 제약조건 문제 해결"""
    print("=== 최적화 제약조건 문제 해결 시작 ===")
    
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
    
    # 1. 선로 용량 증대 (전력 부족 지역 연결 선로)
    print("\n1. 선로 용량 증대")
    lines_df = input_data['lines']
    
    # 부족 지역 목록
    deficit_regions = ['SEL', 'DGU', 'GWJ', 'DJN']
    
    # 부족 지역과 연결된 선로 용량 증대
    for idx, line in lines_df.iterrows():
        bus0 = str(line['bus0'])
        bus1 = str(line['bus1'])
        
        # 부족 지역과 연결된 선로인지 확인
        is_deficit_connected = False
        for region in deficit_regions:
            if f"{region}_{region}_EL" in [bus0, bus1]:
                is_deficit_connected = True
                break
        
        if is_deficit_connected:
            original_capacity = line['s_nom']
            new_capacity = original_capacity * 3  # 3배 증대
            lines_df.at[idx, 's_nom'] = new_capacity
            print(f"  {line['name']}: {original_capacity:,.0f} -> {new_capacity:,.0f} MVA")
    
    # 2. 부족 지역 발전 용량 증대
    print("\n2. 부족 지역 발전 용량 증대")
    generators_df = input_data['generators']
    
    for idx, gen in generators_df.iterrows():
        gen_name = str(gen['name'])
        
        # 부족 지역의 발전기인지 확인
        for region in deficit_regions:
            if gen_name.startswith(f"{region}_"):
                # LNG 발전기 용량 증대 (유연성 확보)
                if 'LNG' in gen_name:
                    original_capacity = gen['p_nom']
                    new_capacity = original_capacity * 2  # 2배 증대
                    generators_df.at[idx, 'p_nom'] = new_capacity
                    print(f"  {gen_name}: {original_capacity:,.0f} -> {new_capacity:,.0f} MW")
                break
    
    # 3. 저장장치 용량 증대 (부족 지역)
    print("\n3. 저장장치 용량 증대")
    stores_df = input_data['stores']
    
    for idx, store in stores_df.iterrows():
        store_name = str(store['name'])
        
        # 부족 지역의 ESS인지 확인
        for region in deficit_regions:
            if store_name.startswith(f"{region}_") and 'ESS' in store_name:
                original_capacity = store['e_nom']
                new_capacity = original_capacity * 2  # 2배 증대
                stores_df.at[idx, 'e_nom'] = new_capacity
                print(f"  {store_name}: {original_capacity:,.0f} -> {new_capacity:,.0f} MWh")
                break
    
    # 4. 발전기 운전 제약 완화
    print("\n4. 발전기 운전 제약 완화")
    
    # p_nom_extendable을 True로 설정 (용량 확장 허용)
    extendable_count = 0
    for idx, gen in generators_df.iterrows():
        if 'p_nom_extendable' not in generators_df.columns:
            generators_df['p_nom_extendable'] = False
        
        # 재생에너지와 LNG 발전기는 확장 가능하도록 설정
        gen_name = str(gen['name'])
        if any(tech in gen_name for tech in ['PV', 'WT', 'LNG']):
            generators_df.at[idx, 'p_nom_extendable'] = True
            extendable_count += 1
    
    print(f"  확장 가능한 발전기: {extendable_count}개")
    
    # 5. 부하 패턴 평활화 (급격한 변동 완화)
    print("\n5. 부하 패턴 평활화")
    
    if 'load_patterns' in input_data:
        load_patterns_df = input_data['load_patterns']
        
        # 부족 지역의 부하 패턴 평활화
        for region in deficit_regions:
            el_col = f"{region}_EL"
            if el_col in load_patterns_df.columns:
                original_pattern = load_patterns_df[el_col].values
                
                # 이동평균으로 평활화 (3시간 윈도우)
                smoothed_pattern = pd.Series(original_pattern).rolling(window=3, center=True).mean()
                smoothed_pattern = smoothed_pattern.fillna(method='bfill').fillna(method='ffill')
                
                load_patterns_df[el_col] = smoothed_pattern.values
                print(f"  {el_col} 패턴 평활화 완료")
    
    # 6. 최소 발전량 제약 완화
    print("\n6. 최소 발전량 제약 완화")
    
    # p_min_pu 컬럼 추가 (없는 경우)
    if 'p_min_pu' not in generators_df.columns:
        generators_df['p_min_pu'] = 0.0
    
    # 석탄 발전기의 최소 발전량 제약 완화
    for idx, gen in generators_df.iterrows():
        gen_name = str(gen['name'])
        if 'Coal' in gen_name:
            generators_df.at[idx, 'p_min_pu'] = 0.0  # 최소 발전량 0%로 설정
    
    print("  석탄 발전기 최소 발전량 제약 완화 완료")
    
    # 7. 수정된 데이터 저장
    print("\n7. 수정된 데이터 저장")
    
    input_data['lines'] = lines_df
    input_data['generators'] = generators_df
    input_data['stores'] = stores_df
    
    with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
        for sheet_name, df in input_data.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print("수정된 데이터가 저장되었습니다.")
    
    # 8. 수정 사항 요약
    print("\n8. 수정 사항 요약:")
    print("  ✓ 부족 지역 연결 선로 용량 3배 증대")
    print("  ✓ 부족 지역 LNG 발전 용량 2배 증대") 
    print("  ✓ 부족 지역 ESS 용량 2배 증대")
    print("  ✓ 재생에너지/LNG 발전기 확장 허용")
    print("  ✓ 부하 패턴 평활화")
    print("  ✓ 석탄 발전기 최소 발전량 제약 완화")
    
    return True

if __name__ == "__main__":
    fix_optimization_constraints() 