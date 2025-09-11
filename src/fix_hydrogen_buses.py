import pandas as pd
import os
import shutil
from datetime import datetime

def fix_hydrogen_buses():
    """
    SEL_Hydrogen 및 JBD_Hydrogen 버스 문제를 해결하는 함수
    1. 누락된 수소 버스 추가
    2. 수소 버스와 전해조 간 연결 확인
    """
    print("수소 버스 문제 수정 중...")
    
    # 파일 경로
    input_file = 'integrated_input_data.xlsx'
    
    # 백업 파일 생성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = f'integrated_input_data_backup_{timestamp}.xlsx'
    
    if not os.path.exists(input_file):
        print(f"오류: {input_file} 파일을 찾을 수 없습니다.")
        return False
    
    # 백업 생성
    shutil.copy2(input_file, backup_file)
    print(f"원본 파일을 {backup_file}로 백업했습니다.")
    
    try:
        # 엑셀 파일 읽기
        buses = pd.read_excel(input_file, sheet_name='buses')
        links = pd.read_excel(input_file, sheet_name='links')
        
        # 누락된 수소 버스 확인 후 추가
        hydrogen_buses_to_add = []
        
        if 'SEL_Hydrogen' not in buses['name'].values:
            hydrogen_buses_to_add.append({
                'name': 'SEL_Hydrogen',
                'v_nom': 345,
                'carrier': 'H2',
                'x': buses.loc[buses['name'] == 'SEL_EL', 'x'].values[0],
                'y': buses.loc[buses['name'] == 'SEL_EL', 'y'].values[0]
            })
            
        if 'JBD_Hydrogen' not in buses['name'].values:
            hydrogen_buses_to_add.append({
                'name': 'JBD_Hydrogen',
                'v_nom': 345,
                'carrier': 'H2',
                'x': buses.loc[buses['name'] == 'JBD_EL', 'x'].values[0],
                'y': buses.loc[buses['name'] == 'JBD_EL', 'y'].values[0]
            })
        
        # 수소 버스 추가
        if hydrogen_buses_to_add:
            buses = pd.concat([buses, pd.DataFrame(hydrogen_buses_to_add)], ignore_index=True)
            print(f"{len(hydrogen_buses_to_add)}개의 수소 버스를 추가했습니다.")
        
        # 엑셀 파일에 저장
        with pd.ExcelWriter(input_file) as writer:
            buses.to_excel(writer, sheet_name='buses', index=False)
            
            # 다른 시트 복사
            sheets = ['generators', 'loads', 'lines', 'stores', 'links']
            for sheet in sheets:
                try:
                    df = pd.read_excel(backup_file, sheet_name=sheet)
                    if sheet == 'links':
                        # SEL_Electrolyser 링크 검사 및 수정
                        sel_idx = df.loc[df['name'] == 'SEL_Electrolyser'].index
                        if len(sel_idx) > 0:
                            print(f"SEL_Electrolyser 링크를 수정합니다: bus1을 'SEL_Hydrogen'으로 설정")
                            df.loc[sel_idx, 'bus1'] = 'SEL_Hydrogen'
                        
                        # JBD_Electrolyser 링크 검사 및 수정
                        jbd_idx = df.loc[df['name'] == 'JBD_Electrolyser'].index
                        if len(jbd_idx) > 0:
                            print(f"JBD_Electrolyser 링크를 수정합니다: bus1을 'JBD_Hydrogen'으로 설정")
                            df.loc[jbd_idx, 'bus1'] = 'JBD_Hydrogen'
                    
                    df.to_excel(writer, sheet_name=sheet, index=False)
                except Exception as e:
                    print(f"시트 '{sheet}' 처리 중 오류 발생: {e}")
        
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
    fix_hydrogen_buses() 