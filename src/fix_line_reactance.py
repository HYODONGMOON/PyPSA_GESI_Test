import pandas as pd
import os

# 입력 및 출력 파일 경로
input_file = 'integrated_input_data.xlsx'
output_file = 'integrated_input_data_fixed_line.xlsx'
backup_file = 'integrated_input_data_backup_line.xlsx'

def fix_line_reactance():
    print("라인 리액턴스 문제 해결 중...")
    
    # 백업 생성
    if os.path.exists(input_file):
        try:
            # 원본 파일 백업 및 사용
            import shutil
            shutil.copy2(input_file, backup_file)
            print(f"원본 파일을 {backup_file}으로 백업했습니다.")
            
            # 원본 파일에서 데이터 로드
            with pd.ExcelFile(input_file) as xls:
                # 모든 시트 로드
                data = {}
                for sheet_name in xls.sheet_names:
                    data[sheet_name] = pd.read_excel(xls, sheet_name)
            
            # JJD_JND 라인의 리액턴스(x) 값 수정
            if 'lines' in data:
                # 문제가 있는 라인 선택
                mask = data['lines']['name'] == 'JJD_JND'
                if mask.any():
                    # 리액턴스 값이 0인지 확인
                    if data['lines'].loc[mask, 'x'].values[0] == 0:
                        # 적절한 리액턴스 값 설정 (r 값의 10배 정도로 설정, 일반적인 AC 라인에서는 x가 r보다 크기 때문)
                        r_value = data['lines'].loc[mask, 'r'].values[0]
                        new_x_value = r_value * 10  # r 값의 10배 정도로 설정
                        
                        # 수정 전 값 저장
                        old_x_value = data['lines'].loc[mask, 'x'].values[0]
                        
                        # x 값 업데이트
                        data['lines'].loc[mask, 'x'] = new_x_value
                        print(f"JJD_JND 라인의 x 값을 {old_x_value}에서 {new_x_value}로 변경했습니다.")
                    else:
                        print(f"JJD_JND 라인의 x 값은 이미 0이 아닙니다. (현재 값: {data['lines'].loc[mask, 'x'].values[0]})")
                else:
                    print("라인 'JJD_JND'를 찾을 수 없습니다.")
            else:
                print("lines 시트가 존재하지 않습니다.")
            
            # 결과 저장
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for sheet_name, df in data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"수정된 데이터가 {output_file}에 저장되었습니다.")
        except Exception as e:
            print(f"오류 발생: {e}")
    else:
        print(f"오류: {input_file} 파일이 존재하지 않습니다.")

if __name__ == "__main__":
    fix_line_reactance() 