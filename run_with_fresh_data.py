import os
import subprocess

# integrated_input_data.xlsx 삭제
integrated_file = 'integrated_input_data.xlsx'
if os.path.exists(integrated_file):
    os.remove(integrated_file)
    print(f"[OK] {integrated_file} 삭제 완료")
else:
    print(f"[INFO] {integrated_file} 없음 (이미 삭제됨)")

# PyPSA_GUI.py 실행
print("\n" + "=" * 80)
print("PyPSA_GUI.py 실행 시작...")
print("=" * 80 + "\n")

subprocess.run(['python', 'PyPSA_GUI.py'])
