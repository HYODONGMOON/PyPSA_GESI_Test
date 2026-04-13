import pandas as pd

print('='*70)
print('지역간 연결 시트 확인')
print('='*70)

try:
    df = pd.read_excel('interface.xlsx', sheet_name='지역간 연결')
    print(f'\n지역간 연결: {len(df)}개')
    print(f'\n컬럼: {list(df.columns)}')
    print(f'\n전체 데이터:')
    print(df.to_string(index=False))
except Exception as e:
    print(f'\n오류: {e}')
    
    # 시트 목록 확인
    print('\n\n모든 시트 목록:')
    xl = pd.ExcelFile('interface.xlsx')
    for sheet in xl.sheet_names:
        print(f'  - {sheet}')

print('\n' + '='*70)
