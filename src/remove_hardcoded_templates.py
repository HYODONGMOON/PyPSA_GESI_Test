import os
import shutil
import re
from datetime import datetime

def remove_hardcoded_templates():
    """
    regional_data_manager.py 파일에서 하드코딩된 REGION_TEMPLATES를 제거합니다.
    """
    print("하드코딩된 템플릿 제거 중...")
    
    # 파일 경로
    target_file = 'regional_data_manager.py'
    backup_file = f'regional_data_manager_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.py'
    
    if not os.path.exists(target_file):
        print(f"오류: {target_file} 파일을 찾을 수 없습니다.")
        return False
    
    # 백업 생성
    shutil.copy2(target_file, backup_file)
    print(f"원본 파일을 {backup_file}로 백업했습니다.")
    
    # 파일 내용 읽기
    with open(target_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # REGION_TEMPLATES 변수 찾기 및 수정
    region_templates_pattern = r'# 지역별 기본 데이터 템플릿\s*REGION_TEMPLATES\s*=\s*\{[^}]*\}'
    
    if re.search(region_templates_pattern, content, re.DOTALL):
        # 하드코딩된 템플릿을 빈 딕셔너리로 대체
        replacement = '# 지역별 기본 데이터 템플릿\nREGION_TEMPLATES = {}'
        content = re.sub(region_templates_pattern, replacement, content, flags=re.DOTALL)
        print("하드코딩된 템플릿이 성공적으로 제거되었습니다.")
    else:
        print("경고: REGION_TEMPLATES 변수를 찾을 수 없습니다.")
    
    # 수정 내용 저장
    with open(target_file, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"\n{target_file} 파일이 성공적으로 수정되었습니다.")
    print("이제 엑셀 템플릿에서 정의되지 않은 요소는 자동으로 추가되지 않습니다.")
    print("변경사항을 적용하려면 PyPSA_GUI.py를 다시 실행하세요.")
    
    # 추가: integrated_input_data.xlsx 파일이 존재하면 삭제하여 새로 생성되도록 함
    integrated_file = 'integrated_input_data.xlsx'
    if os.path.exists(integrated_file):
        try:
            # 백업 후 삭제
            integrated_backup = f'integrated_input_data_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
            shutil.copy2(integrated_file, integrated_backup)
            os.remove(integrated_file)
            print(f"\n{integrated_file} 파일이 삭제되었습니다.")
            print(f"백업 파일: {integrated_backup}")
            print("PyPSA_GUI.py 실행 시 새로운 통합 데이터가 생성됩니다.")
        except Exception as e:
            print(f"파일 삭제 중 오류 발생: {e}")
    
    return True

if __name__ == "__main__":
    remove_hardcoded_templates() 