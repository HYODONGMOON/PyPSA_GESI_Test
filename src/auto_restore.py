import os
import shutil
import pandas as pd
from datetime import datetime

def auto_restore():
    print("가장 적합한 백업에서 파일 복원 중...")
    
    # 원본 파일 경로
    original_file = 'integrated_input_data.xlsx'
    
    # 백업 파일 목록 수집
    backup_files = [f for f in os.listdir('.') if f.startswith('integrated_input_data_') and f.endswith('.xlsx')]
    
    if not backup_files:
        print("사용 가능한 백업 파일이 없습니다.")
        return
    
    # 백업 파일 정보 수집 (크기 기준으로 정렬)
    backup_info = []
    for backup in backup_files:
        try:
            file_size = os.path.getsize(backup)
            file_time = datetime.fromtimestamp(os.path.getmtime(backup))
            
            # 시트 정보 가져오기
            try:
                with pd.ExcelFile(backup) as xls:
                    sheet_count = len(xls.sheet_names)
                    sheets = xls.sheet_names
            except Exception as e:
                sheet_count = 0
                sheets = []
            
            backup_info.append({
                'filename': backup,
                'size': file_size,
                'time': file_time,
                'sheet_count': sheet_count,
                'sheets': sheets
            })
        except Exception as e:
            print(f"파일 {backup} 정보 읽기 오류: {str(e)}")
    
    # 크기 순으로 정렬 (가장 큰 파일이 먼저 오도록)
    backup_info.sort(key=lambda x: x['size'], reverse=True)
    
    print("사용 가능한 백업 파일 (크기 순):")
    for i, info in enumerate(backup_info):
        sheets_info = ', '.join(info['sheets'][:3]) + (', ...' if info['sheet_count'] > 3 else '')
        print(f"{i+1}. {info['filename']} (크기: {info['size']:,} 바이트, 수정일: {info['time'].strftime('%Y-%m-%d %H:%M:%S')}, 시트: {sheets_info})")
    
    # 가장 큰 파일 선택 (일반적으로 가장 완전한 백업일 가능성이 높음)
    if backup_info:
        selected_backup = backup_info[0]['filename']
        
        # 현재 파일 백업
        if os.path.exists(original_file):
            current_backup = f"integrated_input_data_current_before_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            shutil.copy2(original_file, current_backup)
            print(f"\n현재 파일을 {current_backup}으로 백업했습니다.")
        
        # 백업에서 복원
        shutil.copy2(selected_backup, original_file)
        print(f"{selected_backup}에서 {original_file}로 복원이 완료되었습니다.")
        
        # 복원된 파일 정보 확인
        try:
            with pd.ExcelFile(original_file) as xls:
                sheet_names = xls.sheet_names
                print(f"\n복원된 파일의 시트 목록: {', '.join(sheet_names)}")
                
                # 기본 항목 몇 개 확인
                if 'loads' in sheet_names:
                    loads_df = pd.read_excel(xls, 'loads')
                    print(f"로드 항목 수: {len(loads_df)}")
                    
                    # JBD와 SEL 수소 관련 로드 확인
                    problematic_loads = loads_df[loads_df['name'].isin(['JBD_H2_Demand', 'JBD_Demand1', 'SEL_Demand1', 'SEL_H2_Demand'])]
                    if not problematic_loads.empty:
                        print(f"문제가 될 수 있는 로드 항목 {len(problematic_loads)}개 발견:")
                        for _, row in problematic_loads.iterrows():
                            print(f"- {row['name']} (버스: {row['bus']}, p_set: {row['p_set']})")
                    else:
                        print("문제가 될 수 있는 로드 항목이 발견되지 않았습니다.")
                    
                if 'lines' in sheet_names:
                    lines_df = pd.read_excel(xls, 'lines')
                    print(f"라인 항목 수: {len(lines_df)}")
                    
                    # JJD_JND 라인 리액턴스 값 확인
                    jjd_jnd = lines_df[lines_df['name'] == 'JJD_JND']
                    if not jjd_jnd.empty:
                        x_value = jjd_jnd['x'].values[0]
                        r_value = jjd_jnd['r'].values[0]
                        print(f"JJD_JND 라인의 x 값: {x_value}, r 값: {r_value}")
                        if x_value == 0:
                            print("주의: JJD_JND 라인의 x 값이 0입니다. 이는 문제가 될 수 있습니다.")
                
                if 'buses' in sheet_names:
                    buses_df = pd.read_excel(xls, 'buses')
                    print(f"버스 항목 수: {len(buses_df)}")
                    
                    # 수소 버스 확인
                    h2_buses = buses_df[buses_df['name'].isin(['JBD_Hydrogen', 'SEL_Hydrogen'])]
                    if not h2_buses.empty:
                        print(f"수소 버스 {len(h2_buses)}개 발견:")
                        for _, row in h2_buses.iterrows():
                            print(f"- {row['name']}")
        except Exception as e:
            print(f"복원된 파일 정보 확인 중 오류 발생: {str(e)}")
    else:
        print("적합한 백업 파일을 찾을 수 없습니다.")

if __name__ == "__main__":
    auto_restore() 