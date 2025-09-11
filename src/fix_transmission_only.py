import pandas as pd
import numpy as np

def fix_transmission_only():
    """송전선로 조건만 강화"""
    print("=== 송전선로 조건 강화 시작 ===")
    
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
    
    # 현재 선로 상태 확인
    lines_df = input_data['lines']
    print(f"\n현재 선로 수: {len(lines_df)}")
    print("현재 선로 용량:")
    for _, line in lines_df.iterrows():
        print(f"  {line['name']}: {line['s_nom']:,.0f} MVA")
    
    # 1. 모든 선로 용량 대폭 증대 (5배)
    print(f"\n1. 모든 선로 용량 5배 증대")
    for idx, line in lines_df.iterrows():
        original_capacity = line['s_nom']
        new_capacity = original_capacity * 5  # 5배 증대
        lines_df.at[idx, 's_nom'] = new_capacity
        print(f"  {line['name']}: {original_capacity:,.0f} -> {new_capacity:,.0f} MVA")
    
    # 2. 선로 저항 감소 (전력 손실 최소화)
    print(f"\n2. 선로 저항 50% 감소")
    for idx, line in lines_df.iterrows():
        if 'r' in lines_df.columns:
            original_r = line['r']
            new_r = original_r * 0.5  # 50% 감소
            lines_df.at[idx, 'r'] = new_r
            print(f"  {line['name']}: r = {original_r:.4f} -> {new_r:.4f}")
    
    # 3. 선로 리액턴스 감소 (전력 전송 효율 향상)
    print(f"\n3. 선로 리액턴스 50% 감소")
    for idx, line in lines_df.iterrows():
        if 'x' in lines_df.columns:
            original_x = line['x']
            new_x = original_x * 0.5  # 50% 감소
            lines_df.at[idx, 'x'] = new_x
            print(f"  {line['name']}: x = {original_x:.4f} -> {new_x:.4f}")
    
    # 4. 선로 확장 가능성 추가
    print(f"\n4. 선로 확장 가능성 추가")
    if 's_nom_extendable' not in lines_df.columns:
        lines_df['s_nom_extendable'] = True
        print("  모든 선로에 확장 가능 옵션 추가")
    
    # 5. 선로 최대 용량 설정 (현재 용량의 2배까지 확장 가능)
    print(f"\n5. 선로 최대 용량 설정")
    if 's_nom_max' not in lines_df.columns:
        lines_df['s_nom_max'] = lines_df['s_nom'] * 2  # 현재 용량의 2배까지
        print("  선로 최대 용량을 현재 용량의 2배로 설정")
    
    # 6. 부족 지역 연결 선로 추가 강화
    print(f"\n6. 부족 지역 연결 선로 추가 강화")
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
            # 부족 지역 연결 선로는 추가로 2배 더 증대
            current_capacity = lines_df.at[idx, 's_nom']
            enhanced_capacity = current_capacity * 2
            lines_df.at[idx, 's_nom'] = enhanced_capacity
            
            # 최대 용량도 증대
            lines_df.at[idx, 's_nom_max'] = enhanced_capacity * 2
            
            print(f"  {line['name']} (부족지역 연결): {current_capacity:,.0f} -> {enhanced_capacity:,.0f} MVA")
    
    # 7. 수정된 데이터 저장
    print(f"\n7. 수정된 데이터 저장")
    input_data['lines'] = lines_df
    
    with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
        for sheet_name, df in input_data.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print("수정된 데이터가 저장되었습니다.")
    
    # 8. 수정 사항 요약
    print(f"\n8. 수정 사항 요약:")
    print("  ✓ 모든 선로 용량 5배 증대")
    print("  ✓ 선로 저항 50% 감소 (전력 손실 최소화)")
    print("  ✓ 선로 리액턴스 50% 감소 (전송 효율 향상)")
    print("  ✓ 모든 선로 확장 가능 설정")
    print("  ✓ 선로 최대 용량 설정 (현재의 2배)")
    print("  ✓ 부족 지역 연결 선로 추가 2배 강화")
    
    # 최종 선로 용량 확인
    print(f"\n9. 최종 선로 용량:")
    total_capacity = lines_df['s_nom'].sum()
    print(f"  총 선로 용량: {total_capacity:,.0f} MVA")
    
    for _, line in lines_df.iterrows():
        print(f"  {line['name']}: {line['s_nom']:,.0f} MVA (최대: {line['s_nom_max']:,.0f} MVA)")
    
    return True

if __name__ == "__main__":
    fix_transmission_only() 