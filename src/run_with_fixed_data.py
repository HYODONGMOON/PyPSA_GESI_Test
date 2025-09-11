import os
import shutil
import subprocess

# 파일 경로
original_file = 'integrated_input_data.xlsx'
fixed_file = 'integrated_input_data_fixed.xlsx'
backup_file = 'integrated_input_data_backup.xlsx'

def run_with_fixed_data():
    try:
        # 원본 파일 백업
        if os.path.exists(original_file):
            shutil.copy2(original_file, backup_file)
            print(f"원본 파일을 {backup_file}으로 백업했습니다.")
        
        # 수정된 파일을 원본 파일명으로 복사
        shutil.copy2(fixed_file, original_file)
        print(f"수정된 파일을 {original_file}로 복사했습니다.")
        
        # PyPSA_GUI.py 실행
        print("PyPSA_GUI.py 실행 중...")
        subprocess.run(['python', 'PyPSA_GUI.py'], check=True)
        
        print("실행이 완료되었습니다!")
    except Exception as e:
        print(f"오류 발생: {e}")
        # 오류 발생 시 백업 파일 복원
        if os.path.exists(backup_file):
            shutil.copy2(backup_file, original_file)
            print(f"오류가 발생하여 {backup_file}에서 원본 파일을 복원했습니다.")

if __name__ == "__main__":
    run_with_fixed_data() 