import pandas as pd
import os
import shutil
from datetime import datetime

def fix_connections_auto():
    """
    regional_input_template.xlsx 파일의 '지역간 연결' 시트에서 
    문제가 있는 버스 이름을 자동으로 수정합니다.
    """
    print("지역간 연결 자동 수정 중...")
    
    # 파일 경로
    input_file = 'regional_input_template.xlsx'
    backup_file = f'regional_input_template_backup_conn_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    
    if not os.path.exists(input_file):
        print(f"오류: {input_file} 파일을 찾을 수 없습니다.")
        return False
    
    try:
        # 백업 생성
        shutil.copy2(input_file, backup_file)
        print(f"원본 파일을 {backup_file}로 백업했습니다.")
        
        # 시트 이름 확인
        xls = pd.ExcelFile(input_file)
        print(f"엑셀 파일의 시트 목록: {xls.sheet_names}")
        
        connection_sheet = None
        for sheet in xls.sheet_names:
            if '지역간 연결' in sheet or '연결' in sheet or 'connection' in sheet.lower():
                connection_sheet = sheet
                break
        
        if not connection_sheet:
            print("'지역간 연결' 시트를 찾을 수 없습니다.")
            return False
        
        print(f"지역간 연결 시트: '{connection_sheet}'")
        
        # 지역간 연결 데이터 읽기
        connections_df = pd.read_excel(input_file, sheet_name=connection_sheet)
        
        # 열 이름 확인
        print(f"연결 시트의 열: {connections_df.columns.tolist()}")
        
        # bus0와 bus1 열 찾기
        bus0_col = None
        bus1_col = None
        name_col = None
        
        for col in connections_df.columns:
            if 'bus0' in col.lower() or '시작 버스' in col:
                bus0_col = col
            elif 'bus1' in col.lower() or '도착 버스' in col:
                bus1_col = col
            elif 'name' in col.lower() or '이름' in col:
                name_col = col
        
        if not (bus0_col and bus1_col):
            print("버스 연결 정보(bus0, bus1)를 찾을 수 없습니다.")
            return False
        
        print(f"시작 버스 열: '{bus0_col}', 도착 버스 열: '{bus1_col}', 이름 열: '{name_col}'")
        
        # 버스 이름에 _EL 접미사 추가
        modified_count = 0
        for i, row in connections_df.iterrows():
            bus0 = str(row[bus0_col])
            bus1 = str(row[bus1_col])
            
            # 시작 버스 이름 수정
            if not bus0.endswith('_EL') and not bus0.endswith('_Main_EL'):
                # EL이 포함되어 있는지 확인
                if '_EL' not in bus0:
                    new_bus0 = f"{bus0}_EL"
                    connections_df.at[i, bus0_col] = new_bus0
                    print(f"버스 이름 수정: {bus0} -> {new_bus0}")
                    modified_count += 1
            
            # 도착 버스 이름 수정
            if not bus1.endswith('_EL') and not bus1.endswith('_Main_EL'):
                # EL이 포함되어 있는지 확인
                if '_EL' not in bus1:
                    new_bus1 = f"{bus1}_EL"
                    connections_df.at[i, bus1_col] = new_bus1
                    print(f"버스 이름 수정: {bus1} -> {new_bus1}")
                    modified_count += 1
        
        if modified_count == 0:
            print("수정할 버스 이름이 없습니다.")
            return True
        
        # 수정된 데이터 저장
        with pd.ExcelWriter(input_file, engine='openpyxl') as writer:
            for sheet_name in xls.sheet_names:
                if sheet_name == connection_sheet:
                    connections_df.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    pd.read_excel(input_file, sheet_name=sheet_name).to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n{modified_count}개의 버스 이름이 수정되었습니다.")
        print(f"변경사항이 {input_file}에 저장되었습니다.")
        
        # 기존 통합 데이터 파일 삭제
        integrated_file = 'integrated_input_data.xlsx'
        if os.path.exists(integrated_file):
            os.remove(integrated_file)
            print(f"{integrated_file} 파일이 삭제되었습니다.")
            print("PyPSA_GUI.py 실행 시 새로운 통합 데이터가 생성됩니다.")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    fix_connections_auto() 