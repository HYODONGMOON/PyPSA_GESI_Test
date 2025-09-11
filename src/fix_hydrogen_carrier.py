import pandas as pd
import os
import shutil
from datetime import datetime

def fix_hydrogen_carrier():
    """
    수소 버스의 carrier 값을 수정하는 함수
    carrier 값을 H2에서 AC로 변경
    """
    print("수소 버스 carrier 값 수정 중...")
    
    # 파일 경로
    input_file = 'integrated_input_data.xlsx'
    
    # 백업 파일 생성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = f'integrated_input_data_backup_carrier_{timestamp}.xlsx'
    
    if not os.path.exists(input_file):
        print(f"오류: {input_file} 파일을 찾을 수 없습니다.")
        return False
    
    # 백업 생성
    shutil.copy2(input_file, backup_file)
    print(f"원본 파일을 {backup_file}로 백업했습니다.")
    
    try:
        # 엑셀 파일 읽기
        with pd.ExcelFile(input_file) as xls:
            buses = pd.read_excel(xls, sheet_name='buses')
            
            # 다른 시트도 로드
            generators = pd.read_excel(xls, 'generators')
            loads = pd.read_excel(xls, 'loads')
            lines = pd.read_excel(xls, 'lines')
            stores = pd.read_excel(xls, 'stores')
            links = pd.read_excel(xls, 'links')
        
        # 현재 carrier 값 확인
        print("현재 정의된 carrier 값:")
        carrier_values = buses['carrier'].unique()
        for carrier in carrier_values:
            count = buses[buses['carrier'] == carrier].shape[0]
            print(f"  - {carrier}: {count}개")
            
        # 일반적으로 시스템에서 사용되는 carrier 값 확인
        hydrogen_buses = buses[buses['name'].str.contains('Hydrogen')]
        
        if not hydrogen_buses.empty:
            print(f"수소 버스 수: {hydrogen_buses.shape[0]}개")
            
            # 'H2' carrier 값을 'AC'로 변경
            buses.loc[buses['carrier'] == 'H2', 'carrier'] = 'AC'
            print("'H2' carrier를 'AC'로 변경했습니다.")
            
            # 또는 수소 이름을 포함하는 버스의 carrier 값을 변경
            for idx, row in hydrogen_buses.iterrows():
                old_carrier = row['carrier']
                buses.loc[idx, 'carrier'] = 'AC'
                print(f"버스 '{row['name']}'의 carrier를 '{old_carrier}'에서 'AC'로 변경했습니다.")
        
        # 변경된 carrier 값 확인
        print("\n변경 후 정의된 carrier 값:")
        carrier_values = buses['carrier'].unique()
        for carrier in carrier_values:
            count = buses[buses['carrier'] == carrier].shape[0]
            print(f"  - {carrier}: {count}개")
        
        # 엑셀 파일에 저장
        with pd.ExcelWriter(input_file) as writer:
            buses.to_excel(writer, sheet_name='buses', index=False)
            generators.to_excel(writer, sheet_name='generators', index=False)
            loads.to_excel(writer, sheet_name='loads', index=False)
            lines.to_excel(writer, sheet_name='lines', index=False)
            stores.to_excel(writer, sheet_name='stores', index=False)
            links.to_excel(writer, sheet_name='links', index=False)
        
        print(f"\n{input_file} 파일이 성공적으로 수정되었습니다.")
        print("이제 PyPSA_GUI.py를 다시 실행해보세요.")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        
        # 백업에서 복원
        print(f"백업에서 복원 중...")
        shutil.copy2(backup_file, input_file)
        print(f"원본 파일이 백업에서 복원되었습니다.")
        
        return False

if __name__ == "__main__":
    fix_hydrogen_carrier() 