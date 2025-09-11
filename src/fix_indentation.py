import os
import shutil
from datetime import datetime

def fix_indentation():
    """
    regional_data_manager.py 파일의 들여쓰기 오류를 수정합니다.
    """
    print("들여쓰기 오류 수정 중...")
    
    # 파일 경로
    target_file = 'regional_data_manager.py'
    backup_file = f'regional_data_manager_backup_indent_{datetime.now().strftime("%Y%m%d_%H%M%S")}.py'
    
    if not os.path.exists(target_file):
        print(f"오류: {target_file} 파일을 찾을 수 없습니다.")
        return False
    
    # 백업 생성
    shutil.copy2(target_file, backup_file)
    print(f"원본 파일을 {backup_file}로 백업했습니다.")
    
    try:
        # 파일 내용 읽기
        with open(target_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        # 올바른 들여쓰기로 새 내용 생성
        new_lines = []
        in_region_templates = False
        
        for line in lines:
            # REGION_TEMPLATES 변수 선언 확인
            if "REGION_TEMPLATES" in line and "=" in line:
                in_region_templates = True
                # 빈 딕셔너리로 수정
                new_lines.append("REGION_TEMPLATES = {}\n")
                continue
            
            # REGION_TEMPLATES 정의가 끝나는 부분 (클래스 정의 시작) 확인
            if in_region_templates and "class RegionalDataManager" in line:
                in_region_templates = False
            
            # REGION_TEMPLATES 정의 내부는 건너뛰기
            if in_region_templates:
                continue
            
            # 다른 모든 라인은 그대로 유지
            new_lines.append(line)
        
        # 수정된 내용 저장
        with open(target_file, 'w', encoding='utf-8') as f:
            f.writelines(new_lines)
        
        print(f"\n{target_file} 파일의 들여쓰기 오류가 수정되었습니다.")
        print("이제 PyPSA_GUI.py를 다시 실행해보세요.")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    fix_indentation() 