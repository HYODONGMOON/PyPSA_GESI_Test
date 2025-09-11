import pandas as pd
import os
from datetime import datetime

def update_integrated_data():
    """regional_input_template.xlsx에서 데이터를 읽어와서 integrated_input_data.xlsx를 업데이트"""
    
    try:
        print("regional_input_template.xlsx에서 데이터 읽는 중...")
        
        # regional_input_template.xlsx 읽기
        template_file = "regional_input_template.xlsx"
        if not os.path.exists(template_file):
            print(f"오류: {template_file} 파일이 존재하지 않습니다.")
            return False
        
        # 통합 시트들 읽기
        sheets_to_copy = [
            '통합_buses', '통합_generators', '통합_loads', 
            '통합_lines', '통합_stores', '통합_links',
            'constraints', 'renewable_patterns', 'load_patterns',
            '시간 설정'
        ]
        
        # 새로운 integrated_input_data.xlsx 생성
        output_file = "integrated_input_data.xlsx"
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            
            # 각 시트 복사
            for sheet_name in sheets_to_copy:
                try:
                    if sheet_name == '시간 설정':
                        # 시간 설정 시트를 timeseries로 변경
                        df = pd.read_excel(template_file, sheet_name=sheet_name)
                        df.to_excel(writer, sheet_name='timeseries', index=False)
                        print(f"시트 복사됨: {sheet_name} -> timeseries")
                    else:
                        # 통합_ 접두사 제거
                        target_sheet = sheet_name.replace('통합_', '')
                        df = pd.read_excel(template_file, sheet_name=sheet_name)
                        df.to_excel(writer, sheet_name=target_sheet, index=False)
                        print(f"시트 복사됨: {sheet_name} -> {target_sheet}")
                        
                except Exception as e:
                    print(f"시트 {sheet_name} 복사 중 오류: {str(e)}")
                    continue
        
        print(f"\n{output_file} 파일이 성공적으로 업데이트되었습니다.")
        
        # 파일 수정 시간 확인
        mod_time = os.path.getmtime(output_file)
        mod_datetime = datetime.fromtimestamp(mod_time)
        print(f"파일 수정 시간: {mod_datetime}")
        
        return True
        
    except Exception as e:
        print(f"데이터 업데이트 중 오류 발생: {str(e)}")
        return False

def check_template_data():
    """regional_input_template.xlsx의 데이터 구조 확인"""
    
    try:
        template_file = "regional_input_template.xlsx"
        xls = pd.ExcelFile(template_file)
        
        print("=== regional_input_template.xlsx 데이터 구조 ===")
        
        # 통합 시트들 확인
        for sheet_name in xls.sheet_names:
            if '통합_' in sheet_name or sheet_name in ['시간 설정', 'constraints', 'renewable_patterns', 'load_patterns']:
                df = pd.read_excel(template_file, sheet_name=sheet_name)
                print(f"\n[{sheet_name}] 시트:")
                print(f"행 수: {len(df)}")
                print(f"컬럼: {list(df.columns)}")
                
                # 처음 몇 행 데이터 확인
                if len(df) > 0:
                    print("첫 3행 데이터:")
                    print(df.head(3))
        
        return True
        
    except Exception as e:
        print(f"템플릿 데이터 확인 중 오류: {str(e)}")
        return False

if __name__ == "__main__":
    print("=== 템플릿 데이터 구조 확인 ===")
    check_template_data()
    
    print("\n=== integrated_input_data.xlsx 업데이트 ===")
    update_integrated_data() 