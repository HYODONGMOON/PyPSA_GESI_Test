import pandas as pd
import os
import shutil

# 파일 경로
input_file = 'regional_input_template.xlsx'
backup_file = 'regional_input_template_backup.xlsx'

def fix_regional_template():
    print("원본 입력 템플릿 파일 수정 중...")
    
    # 백업 생성
    if os.path.exists(input_file):
        try:
            # 원본 파일 백업
            shutil.copy2(input_file, backup_file)
            print(f"원본 파일을 {backup_file}으로 백업했습니다.")
            
            # 원본 파일에서 데이터 로드
            with pd.ExcelFile(input_file) as xls:
                # 모든 시트 이름 확인
                sheet_names = xls.sheet_names
                print(f"파일에 포함된 시트: {', '.join(sheet_names)}")
                
                # 모든 시트 로드
                data = {}
                for sheet_name in sheet_names:
                    data[sheet_name] = pd.read_excel(xls, sheet_name)
            
            # 1. 'loads' 시트에서 지정된 로드 항목 제거
            if 'loads' in data:
                print("\n[1] loads 시트 처리 중...")
                # 제거할 로드 이름 목록
                loads_to_remove = ['JBD_H2_Demand', 'JBD_Demand1', 'SEL_Demand1', 'SEL_H2_Demand']
                
                # 제거하기 전 정보 출력
                removed_loads = data['loads'][data['loads']['name'].isin(loads_to_remove)]
                print(f"제거할 로드 수: {len(removed_loads)}")
                if not removed_loads.empty:
                    for idx, row in removed_loads.iterrows():
                        print(f"제거: {row['name']} (버스: {row['bus']}, p_set: {row['p_set']})")
                    
                    # 로드 제거
                    data['loads'] = data['loads'][~data['loads']['name'].isin(loads_to_remove)]
                    print("지정된 로드 항목 제거 완료")
                else:
                    print("제거할 로드 항목이 없습니다.")
            else:
                print("\n[1] 'loads' 시트가 존재하지 않습니다.")
            
            # 2. 'lines' 시트에서 JJD_JND 라인의 리액턴스(x) 값 수정
            if 'lines' in data:
                print("\n[2] lines 시트 처리 중...")
                # JJD_JND 라인 선택
                mask = data['lines']['name'] == 'JJD_JND'
                if mask.any():
                    # 리액턴스 값 확인
                    x_value = data['lines'].loc[mask, 'x'].values[0]
                    r_value = data['lines'].loc[mask, 'r'].values[0]
                    
                    print(f"JJD_JND 라인의 현재 x 값: {x_value}, r 값: {r_value}")
                    
                    # 리액턴스 값이 0이면 수정
                    if x_value == 0:
                        new_x_value = r_value * 10  # r 값의 10배 정도로 설정
                        data['lines'].loc[mask, 'x'] = new_x_value
                        print(f"JJD_JND 라인의 x 값을 {x_value}에서 {new_x_value}로 변경했습니다.")
                    else:
                        print("JJD_JND 라인의 x 값은 이미 0이 아닙니다. 수정이 필요하지 않습니다.")
                else:
                    print("JJD_JND 라인을 찾을 수 없습니다.")
            else:
                print("\n[2] 'lines' 시트가 존재하지 않습니다.")
            
            # 결과 저장 (원본 파일에 덮어쓰기)
            with pd.ExcelWriter(input_file, engine='openpyxl') as writer:
                for sheet_name, df in data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"\n모든 수정이 완료되었습니다. 원본 입력 템플릿 '{input_file}'이 수정되었습니다.")
            print(f"백업 파일은 '{backup_file}'에 저장되어 있습니다.")
            print("\n이제 PyPSA_GUI.py를 실행하여 수정된 템플릿으로 새로운 네트워크를 생성할 수 있습니다.")
        except Exception as e:
            print(f"오류 발생: {e}")
    else:
        print(f"오류: {input_file} 파일이 존재하지 않습니다.")

if __name__ == "__main__":
    fix_regional_template() 