import os
import shutil
from datetime import datetime

def fix_class_structure():
    """
    regional_data_manager.py 파일의 클래스 구조를 수정하는 함수
    1. __init__ 메서드를 클래스 정의 직후로 이동
    2. 코드 구조 정리
    """
    print("regional_data_manager.py 파일의 클래스 구조 수정 중...")
    
    # 파일 경로
    target_file = 'regional_data_manager.py'
    backup_file = f'regional_data_manager_backup_struct_{datetime.now().strftime("%Y%m%d_%H%M%S")}.py'
    
    if not os.path.exists(target_file):
        print(f"오류: {target_file} 파일을 찾을 수 없습니다.")
        return False
    
    # 백업 생성
    shutil.copy2(target_file, backup_file)
    print(f"원본 파일을 {backup_file}로 백업했습니다.")
    
    try:
        # 파일 내용 읽기
        with open(target_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 클래스 정의와 메서드들 추출
        class_start = content.find("class RegionalDataManager")
        if class_start == -1:
            print("오류: RegionalDataManager 클래스를 찾을 수 없습니다.")
            return False
        
        # 클래스 선언부 추출
        class_declaration = content[class_start:content.find("\n", class_start) + 1]
        
        # __init__ 메서드 추출
        init_start = content.find("def __init__", class_start)
        if init_start == -1:
            print("오류: __init__ 메서드를 찾을 수 없습니다.")
            return False
        
        init_end = content.find("\n\n", init_start)
        if init_end == -1:  # 파일 끝까지 init 메서드가 이어지는 경우
            init_end = len(content)
        
        init_method = content[init_start:init_end]
        
        # 다른 메서드들 추출
        other_methods = []
        methods_start = content.find("def ", class_start)
        
        while methods_start != -1 and methods_start < init_start:
            method_end = content.find("\n\n", methods_start)
            if method_end == -1 or method_end > init_start:
                method_end = init_start
            
            method = content[methods_start:method_end]
            other_methods.append(method)
            
            methods_start = content.find("def ", method_end)
            if methods_start >= init_start or methods_start == -1:
                break
        
        # 클래스 정의 전 코드 추출
        before_class = content[:class_start]
        
        # 클래스 이후 코드 추출 (없을 수 있음)
        after_class = content[init_end:] if init_end < len(content) else ""
        
        # 새로운 클래스 구조 작성
        new_class = f"{class_declaration}\n    {init_method.replace('def __init__', 'def __init__')}\n\n"
        
        for method in other_methods:
            new_class += f"    {method}\n\n"
        
        # 최종 파일 내용 생성
        new_content = before_class + new_class + after_class
        
        # 수정된 내용 저장
        with open(target_file, 'w', encoding='utf-8') as f:
            f.write(new_content)
        
        print(f"\n{target_file} 파일의 클래스 구조가 성공적으로 수정되었습니다.")
        print("이제 PyPSA_GUI.py를 다시 실행해보세요.")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        
        # 백업에서 복원
        print(f"백업에서 복원 중...")
        shutil.copy2(backup_file, target_file)
        print(f"원본 파일이 백업에서 복원되었습니다.")
        
        return False

if __name__ == "__main__":
    fix_class_structure() 