import pandas as pd
import os
import shutil
from datetime import datetime

def fix_connections_header():
    """
    regional_input_template.xlsx 파일의 '지역간 연결' 시트에서 
    올바른 헤더를 인식하고 버스 이름을 수정합니다.
    """
    print("지역간 연결 헤더 및 버스 이름 수정 중...")
    
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
        
        # 전체 시트 목록 확인
        xls = pd.ExcelFile(input_file)
        print(f"엑셀 파일의 시트 목록: {xls.sheet_names}")
        
        # '지역간 연결' 시트 찾기
        if '지역간 연결' not in xls.sheet_names:
            print("'지역간 연결' 시트를 찾을 수 없습니다.")
            return False
        
        # 헤더 없이 전체 데이터 읽기
        raw_df = pd.read_excel(input_file, sheet_name='지역간 연결', header=None)
        
        # 데이터 확인 (처음 10행)
        print("원본 데이터 (처음 10행):")
        print(raw_df.head(10))
        
        # 헤더 행 찾기 - '이름', '시작 지역', '시작 버스' 등이 있는 행
        header_row = None
        for i in range(len(raw_df)):
            if raw_df.iloc[i].astype(str).str.contains('이름').any() and raw_df.iloc[i].astype(str).str.contains('시작 버스').any():
                header_row = i
                break
        
        if header_row is None:
            print("헤더 행을 찾을 수 없습니다.")
            return False
        
        print(f"헤더 행 인덱스: {header_row}")
        
        # 헤더 행을 이용해 데이터 다시 읽기
        connections_df = pd.read_excel(input_file, sheet_name='지역간 연결', header=header_row)
        
        # 열 이름 확인
        print(f"연결 시트의 열: {connections_df.columns.tolist()}")
        
        # 시작 버스, 도착 버스 열 찾기
        start_bus_col = None
        end_bus_col = None
        name_col = None
        
        for col in connections_df.columns:
            if '시작 버스' in str(col):
                start_bus_col = col
            elif '도착 버스' in str(col):
                end_bus_col = col
            elif '이름' in str(col):
                name_col = col
        
        if not (start_bus_col and end_bus_col):
            print("버스 연결 정보(시작 버스, 도착 버스)를 찾을 수 없습니다.")
            return False
        
        print(f"시작 버스 열: '{start_bus_col}', 도착 버스 열: '{end_bus_col}', 이름 열: '{name_col}'")
        
        # 실제 데이터 시작 행 찾기 - NaN 값이 아닌 첫 번째 행
        first_data_row = None
        for i in range(len(connections_df)):
            if pd.notna(connections_df.iloc[i][name_col]) and pd.notna(connections_df.iloc[i][start_bus_col]):
                first_data_row = i
                break
        
        if first_data_row is None:
            print("데이터 행을 찾을 수 없습니다.")
            return False
        
        print(f"첫 번째 데이터 행 인덱스: {first_data_row}")
        
        # 실제 데이터만 추출
        connections_df = connections_df.iloc[first_data_row:].reset_index(drop=True)
        
        # 데이터 확인
        print("\n실제 데이터 (처음 5행):")
        print(connections_df.head())
        
        # 버스 이름에 _EL 접미사 추가
        modified_count = 0
        for i, row in connections_df.iterrows():
            bus0 = str(row[start_bus_col])
            bus1 = str(row[end_bus_col])
            
            # 빈 값 또는 NaN 건너뛰기
            if pd.isna(bus0) or pd.isna(bus1) or bus0 == 'nan' or bus1 == 'nan':
                continue
            
            # 시작 버스 이름 수정
            if not bus0.endswith('_EL') and not bus0.endswith('_Main_EL'):
                # EL이 포함되어 있는지 확인
                if '_EL' not in bus0:
                    new_bus0 = f"{bus0}_EL"
                    connections_df.at[i, start_bus_col] = new_bus0
                    print(f"시작 버스 이름 수정: {bus0} -> {new_bus0}")
                    modified_count += 1
            
            # 도착 버스 이름 수정
            if not bus1.endswith('_EL') and not bus1.endswith('_Main_EL'):
                # EL이 포함되어 있는지 확인
                if '_EL' not in bus1:
                    new_bus1 = f"{bus1}_EL"
                    connections_df.at[i, end_bus_col] = new_bus1
                    print(f"도착 버스 이름 수정: {bus1} -> {new_bus1}")
                    modified_count += 1
        
        if modified_count == 0:
            print("수정할 버스 이름이 없습니다.")
            return True
        
        # 수정된 데이터를 시트에 다시 쓰기
        # 원본 Excel 파일 전체 읽기
        all_data = {}
        for sheet in xls.sheet_names:
            all_data[sheet] = pd.read_excel(input_file, sheet_name=sheet)
        
        # 지역간 연결 시트만 업데이트
        # 헤더가 원본 위치를 유지하도록 처리
        raw_data = pd.read_excel(input_file, sheet_name='지역간 연결', header=None)
        
        # 헤더 행 이후 데이터를 수정된 데이터로 대체
        for i in range(header_row + 1 + first_data_row, min(header_row + 1 + first_data_row + len(connections_df), len(raw_data))):
            if i - (header_row + 1 + first_data_row) < len(connections_df):
                for j, col in enumerate(connections_df.columns):
                    raw_data.iloc[i, j] = connections_df.iloc[i - (header_row + 1 + first_data_row)][col]
        
        # 수정된 데이터 저장
        with pd.ExcelWriter(input_file, engine='openpyxl') as writer:
            for sheet in xls.sheet_names:
                if sheet == '지역간 연결':
                    raw_data.to_excel(writer, sheet_name=sheet, header=False, index=False)
                else:
                    all_data[sheet].to_excel(writer, sheet_name=sheet, index=False)
        
        print(f"\n{modified_count}개의 버스 이름이 수정되었습니다.")
        print(f"변경사항이 {input_file}에 저장되었습니다.")
        
        # 기존 통합 데이터 파일 삭제
        integrated_file = 'integrated_input_data.xlsx'
        if os.path.exists(integrated_file):
            os.remove(integrated_file)
            print(f"{integrated_file} 파일이 삭제되었습니다.")
            print("PyPSA_GUI.py 실행 시 새로운 통합 데이터가 생성됩니다.")
        
        # 간단하게 문제 해결 시도 - 파일 수정이 복잡한 경우
        print("\n복잡한 엑셀 구조로 인해 자동 수정이 어려울 수 있습니다.")
        print("다음 방법을 직접 시도해보세요:")
        print("1. regional_input_template.xlsx 파일을 Excel로 열기")
        print("2. '지역간 연결' 시트의 '시작 버스' 열에서 버스 이름 뒤에 '_EL'이 없는 경우 추가 (예: BSN → BSN_EL)")
        print("3. '도착 버스' 열에서도 동일하게 버스 이름 뒤에 '_EL'이 없는 경우 추가")
        print("4. 파일 저장 후 PyPSA_GUI.py 다시 실행")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    fix_connections_header() 