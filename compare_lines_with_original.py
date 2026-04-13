import pandas as pd
import os

print('='*70)
print('현재 모델 vs 이전 모델 lines 비교')
print('='*70)

# 현재 모델
current_df = pd.read_excel('interface.xlsx', sheet_name='lines')
current_self_loops = current_df[current_df['bus0'] == current_df['bus1']]

print(f'\n[현재 모델]')
print(f'  총 선로: {len(current_df)}개')
print(f'  Self-loop: {len(current_self_loops)}개 ({len(current_self_loops)/len(current_df)*100:.1f}%)')
print(f'  정상 선로: {len(current_df)-len(current_self_loops)}개')

# 이전 모델
original_path = r'C:\Users\Hyodong.Moon\Desktop\HDMOON\python workplace\Pypsa_test_original\interface.xlsx'
if os.path.exists(original_path):
    original_df = pd.read_excel(original_path, sheet_name='lines')
    original_self_loops = original_df[original_df['bus0'] == original_df['bus1']]
    
    print(f'\n[이전 모델]')
    print(f'  총 선로: {len(original_df)}개')
    print(f'  Self-loop: {len(original_self_loops)}개 ({len(original_self_loops)/len(original_df)*100:.1f}%)')
    print(f'  정상 선로: {len(original_df)-len(original_self_loops)}개')
    
    print(f'\n[차이]')
    print(f'  총 선로 증가: {len(current_df) - len(original_df)}개')
    print(f'  Self-loop 증가: {len(current_self_loops) - len(original_self_loops)}개')
else:
    print(f'\n[이전 모델] 파일을 찾을 수 없습니다: {original_path}')

print('\n' + '='*70)
