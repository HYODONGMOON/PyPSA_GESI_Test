import os
import shutil
import re
from datetime import datetime

def fix_template_apply():
    """
    regional_data_manager.py 파일을 수정하여 하드코딩된 템플릿이 적용되지 않도록 합니다.
    """
    print("템플릿 자동 적용 문제 수정 중...")
    
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
    
    # 수정 방법 1: initialize_region 메서드에서 _apply_template 호출을 조건부로 변경
    # 기존 코드:
    # if region_code in self.region_templates:
    #     self._apply_template(region_code)
    
    # 새 코드:
    # if region_code in self.region_templates and not use_excel_template_only:
    #     self._apply_template(region_code)
    
    # 변수 추가를 먼저 수행
    class_init_pattern = r'def __init__\(self, region_selector\):'
    if re.search(class_init_pattern, content):
        replacement = 'def __init__(self, region_selector):\n        """\n        Args:\n            region_selector (RegionalSelector): 지역 선택 객체\n        """\n        self.region_selector = region_selector\n        self.regional_data = {}  # 지역별 데이터 저장\n        self.connections = []    # 지역간 연결 저장\n        self.use_excel_template_only = True  # 엑셀 템플릿만 사용 (하드코딩된 템플릿 무시)\n        \n        # 기본 템플릿 로드'
        content = re.sub(class_init_pattern + r'[\s\S]+?# 기본 템플릿 로드', replacement, content)
    else:
        print("경고: __init__ 메서드를 찾을 수 없습니다.")
    
    # _apply_template 호출 부분 수정
    template_call_pattern = r'if region_code in self\.region_templates:'
    if re.search(template_call_pattern, content):
        replacement = 'if region_code in self.region_templates and not self.use_excel_template_only:'
        content = content.replace('if region_code in self.region_templates:', replacement)
        print("_apply_template 호출 부분을 조건부로 수정했습니다.")
    else:
        print("경고: _apply_template 호출 부분을 찾을 수 없습니다.")
    
    # 수정 내용 저장
    with open(target_file, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"\n{target_file} 파일이 성공적으로 수정되었습니다.")
    print("이제 엑셀 템플릿에서 정의되지 않은 요소는 자동으로 추가되지 않습니다.")
    print("변경사항을 적용하려면 PyPSA_GUI.py를 다시 실행하세요.")
    
    return True

if __name__ == "__main__":
    fix_template_apply() 