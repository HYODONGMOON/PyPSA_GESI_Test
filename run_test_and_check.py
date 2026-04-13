# -*- coding: utf-8 -*-
"""테스트 실행 및 로그 확인"""
import subprocess
import sys

# 테스트 실행
print("테스트 실행 중...")
result = subprocess.run(
    [sys.executable, "test_bus_name_generation.py"],
    capture_output=True,
    text=True,
    encoding='utf-8',
    errors='ignore'
)

# DEBUG 메시지만 필터링
print("\n" + "="*70)
print("DEBUG 메시지:")
print("="*70)
for line in result.stdout.split('\n'):
    if 'DEBUG' in line or 'virtual_buses' in line or 'all_buses' in line:
        print(line)

# 전체 출력이 필요하면 파일에 저장
with open('full_test_output.txt', 'w', encoding='utf-8', errors='ignore') as f:
    f.write(result.stdout)
    f.write("\n\n=== STDERR ===\n\n")
    f.write(result.stderr)

print("\n전체 로그는 full_test_output.txt에 저장되었습니다.")
