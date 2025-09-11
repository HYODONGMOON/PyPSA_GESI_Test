import shutil
import os
from datetime import datetime

def restore_template():
    print("regional_input_template 복원 중...")
    
    # 파일 경로
    original_file = 'regional_input_template.xlsx'
    backup_file = 'regional_input_template_backup.xlsx'
    current_backup = f'regional_input_template_current_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    
    if os.path.exists(backup_file):
        try:
            # 현재 파일 백업
            if os.path.exists(original_file):
                shutil.copy2(original_file, current_backup)
                print(f"현재 템플릿 파일을 {current_backup}으로 백업했습니다.")
            
            # 백업에서 복원
            shutil.copy2(backup_file, original_file)
            print(f"{backup_file}에서 {original_file}로 복원이 완료되었습니다.")
            
            # 파일 크기 확인
            original_size = os.path.getsize(original_file)
            backup_size = os.path.getsize(backup_file)
            print(f"복원된 파일 크기: {original_size:,} 바이트")
            print(f"백업 파일 크기: {backup_size:,} 바이트")
            
            if original_size == backup_size:
                print("파일 크기가 일치합니다. 서식과 하이퍼링크가 성공적으로 복원되었을 가능성이 높습니다.")
            else:
                print("파일 크기가 다릅니다. 복원 과정에서 문제가 발생했을 수 있습니다.")
        except Exception as e:
            print(f"오류 발생: {e}")
    else:
        print(f"오류: 백업 파일 {backup_file}이 존재하지 않습니다.")

if __name__ == "__main__":
    restore_template() 