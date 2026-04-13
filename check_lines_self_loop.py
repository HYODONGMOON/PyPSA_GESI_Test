import pandas as pd

df = pd.read_excel('interface.xlsx', sheet_name='lines')

print('='*70)
print('lines 시트 Self-loop 진단')
print('='*70)

print(f'\n총 선로: {len(df)}개')

print(f'\n컬럼: {list(df.columns)}')

print(f'\n처음 20개:')
print(df[['name', 'bus0', 'bus1', 's_nom']].head(20).to_string(index=False))

# Self-loop 확인
self_loops = df[df['bus0'] == df['bus1']]
print(f'\n\nSelf-loop 선로: {len(self_loops)}개 ({len(self_loops)/len(df)*100:.1f}%)')

if len(self_loops) > 0:
    print('\nSelf-loop 샘플 (처음 30개):')
    print(self_loops[['name', 'bus0', 'bus1', 's_nom']].head(30).to_string(index=False))

# 정상 선로 확인
normal_lines = df[df['bus0'] != df['bus1']]
print(f'\n\n정상 선로: {len(normal_lines)}개 ({len(normal_lines)/len(df)*100:.1f}%)')

if len(normal_lines) > 0:
    print('\n정상 선로 샘플 (처음 20개):')
    print(normal_lines[['name', 'bus0', 'bus1', 's_nom']].head(20).to_string(index=False))

print('\n' + '='*70)
