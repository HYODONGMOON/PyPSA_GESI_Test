import pandas as pd
import numpy as np

def fix_transmission_only_v2():
    """발전기 용량 원상복구 + 송전선로만 대폭 강화"""
    print("=== 발전기 용량 원상복구 + 송전선로 대폭 강화 ===")
    
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
    
    # 1. 발전기 용량 원상복구
    print(f"\n1. 발전기 용량 원상복구")
    generators_df = input_data['generators']
    
    # 원래 용량으로 복구
    original_capacities = {
        'DGU_LNG': 510,
        'DJN_LNG': 50, 
        'GWJ_LNG': 60,
        'SEL_LNG': 800
    }
    
    for idx, gen in generators_df.iterrows():
        gen_name = gen['name']
        if gen_name in original_capacities:
            original_capacity = original_capacities[gen_name]
            current_capacity = gen['p_nom']
            generators_df.at[idx, 'p_nom'] = original_capacity
            print(f"  {gen_name}: {current_capacity} -> {original_capacity} MW (원상복구)")
    
    # 2. 송전선로 용량 대폭 강화 (10배)
    print(f"\n2. 송전선로 용량 10배 강화")
    lines_df = input_data['lines']
    
    for idx, line in lines_df.iterrows():
        original_capacity = line['s_nom']
        new_capacity = original_capacity * 10  # 10배 증대
        lines_df.at[idx, 's_nom'] = new_capacity
        print(f"  {line['name']}: {original_capacity:,.0f} -> {new_capacity:,.0f} MVA")
    
    # 3. 선로 저항/리액턴스 대폭 감소 (20%로)
    print(f"\n3. 선로 저항/리액턴스 80% 감소")
    for idx, line in lines_df.iterrows():
        if 'r' in lines_df.columns:
            original_r = line['r']
            new_r = original_r * 0.2  # 80% 감소
            lines_df.at[idx, 'r'] = new_r
            print(f"  {line['name']}: r = {original_r:.4f} -> {new_r:.4f}")
        
        if 'x' in lines_df.columns:
            original_x = line['x']
            new_x = original_x * 0.2  # 80% 감소
            lines_df.at[idx, 'x'] = new_x
            print(f"  {line['name']}: x = {original_x:.4f} -> {new_x:.4f}")
    
    # 4. 선로 확장성 대폭 강화
    print(f"\n4. 선로 확장성 대폭 강화")
    lines_df['s_nom_extendable'] = True
    lines_df['s_nom_max'] = lines_df['s_nom'] * 5  # 현재 용량의 5배까지 확장 가능
    print("  모든 선로 확장 가능, 최대 용량 5배 설정")
    
    # 5. 부족 지역 연결 선로 추가 강화 (추가 5배)
    print(f"\n5. 부족 지역 연결 선로 추가 강화")
    deficit_regions = ['SEL', 'DGU', 'GWJ', 'DJN']
    
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
            # 부족 지역 연결 선로는 추가로 5배 더 증대
            current_capacity = lines_df.at[idx, 's_nom']
            enhanced_capacity = current_capacity * 5
            lines_df.at[idx, 's_nom'] = enhanced_capacity
            
            # 최대 용량도 대폭 증대
            lines_df.at[idx, 's_nom_max'] = enhanced_capacity * 10
            
            print(f"  {line['name']} (부족지역 연결): {current_capacity:,.0f} -> {enhanced_capacity:,.0f} MVA")
    
    # 6. 수정된 데이터 저장
    print(f"\n6. 수정된 데이터 저장")
    input_data['generators'] = generators_df
    input_data['lines'] = lines_df
    
    with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
        for sheet_name, df in input_data.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print("수정된 데이터가 저장되었습니다.")
    
    # 7. 수정 사항 요약
    print(f"\n7. 수정 사항 요약:")
    print("  ✓ 발전기 용량 원상복구 (DGU_LNG: 510MW, DJN_LNG: 50MW, GWJ_LNG: 60MW, SEL_LNG: 800MW)")
    print("  ✓ 모든 선로 용량 10배 증대")
    print("  ✓ 선로 저항/리액턴스 80% 감소")
    print("  ✓ 모든 선로 확장 가능 (최대 5배)")
    print("  ✓ 부족 지역 연결 선로 추가 5배 강화")
    
    # 최종 선로 용량 확인
    print(f"\n8. 최종 선로 용량:")
    total_capacity = lines_df['s_nom'].sum()
    print(f"  총 선로 용량: {total_capacity:,.0f} MVA")
    
    # 발전기 용량 확인
    print(f"\n9. 복구된 발전기 용량:")
    for gen_name in original_capacities.keys():
        gen_row = generators_df[generators_df['name'] == gen_name]
        if not gen_row.empty:
            print(f"  {gen_name}: {gen_row.iloc[0]['p_nom']} MW")
    
    return True

if __name__ == "__main__":
    fix_transmission_only_v2() 