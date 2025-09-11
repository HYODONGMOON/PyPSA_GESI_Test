import os
import shutil
import pandas as pd
from datetime import datetime

def restore_from_backup():
    print("백업에서 파일 복원 중...")
    
    # 원본 파일 경로
    original_file = 'integrated_input_data.xlsx'
    
    # 백업 파일 목록 수집
    backup_files = [f for f in os.listdir('.') if f.startswith('integrated_input_data_') and f.endswith('.xlsx')]
    backup_files.sort()  # 파일명 기준 정렬
    
    if not backup_files:
        print("사용 가능한 백업 파일이 없습니다.")
        return
    
    print("사용 가능한 백업 파일:")
    for i, backup in enumerate(backup_files):
        try:
            file_size = os.path.getsize(backup)
            file_time = datetime.fromtimestamp(os.path.getmtime(backup)).strftime('%Y-%m-%d %H:%M:%S')
            
            # 시트 정보 가져오기
            try:
                with pd.ExcelFile(backup) as xls:
                    sheet_count = len(xls.sheet_names)
                    sheet_info = ', '.join(xls.sheet_names[:3]) + (', ...' if sheet_count > 3 else '')
            except Exception as e:
                sheet_info = f"시트 정보 읽기 오류: {str(e)}"
            
            print(f"{i+1}. {backup} (크기: {file_size:,} 바이트, 수정일: {file_time}, 시트: {sheet_info})")
        except Exception as e:
            print(f"{i+1}. {backup} (정보 읽기 오류: {str(e)})")
    
    # 선택을 위한 입력
    print("\n원하는 백업 파일 번호를 입력하세요 (또는 q로 종료): ", end='')
    choice = input().strip()
    
    if choice.lower() == 'q':
        print("복원 취소됨.")
        return
    
    try:
        choice_idx = int(choice) - 1
        if 0 <= choice_idx < len(backup_files):
            selected_backup = backup_files[choice_idx]
            
            # 현재 파일 백업
            if os.path.exists(original_file):
                current_backup = f"integrated_input_data_current_before_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                shutil.copy2(original_file, current_backup)
                print(f"현재 파일을 {current_backup}으로 백업했습니다.")
            
            # 백업에서 복원
            shutil.copy2(selected_backup, original_file)
            print(f"{selected_backup}에서 {original_file}로 복원이 완료되었습니다.")
            
            # 복원된 파일 정보 확인
            with pd.ExcelFile(original_file) as xls:
                sheet_names = xls.sheet_names
                print(f"복원된 파일의 시트 목록: {', '.join(sheet_names)}")
                
                # 기본 항목 몇 개 확인
                if 'loads' in sheet_names:
                    loads_df = pd.read_excel(xls, 'loads')
                    print(f"로드 항목 수: {len(loads_df)}")
                    
                if 'lines' in sheet_names:
                    lines_df = pd.read_excel(xls, 'lines')
                    print(f"라인 항목 수: {len(lines_df)}")
                    
                    # JJD_JND 라인 리액턴스 값 확인
                    jjd_jnd = lines_df[lines_df['name'] == 'JJD_JND']
                    if not jjd_jnd.empty:
                        print(f"JJD_JND 라인의 x 값: {jjd_jnd['x'].values[0]}")
                
                if 'buses' in sheet_names:
                    buses_df = pd.read_excel(xls, 'buses')
                    print(f"버스 항목 수: {len(buses_df)}")
        else:
            print("잘못된 번호를 입력했습니다.")
    except ValueError:
        print("유효한 숫자를 입력해주세요.")
    except Exception as e:
        print(f"오류 발생: {str(e)}")

if __name__ == "__main__":
    restore_from_backup() 