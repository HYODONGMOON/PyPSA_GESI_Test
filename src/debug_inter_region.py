import pandas as pd

# 지역간 연결 시트 상세 분석
df = pd.read_excel('regional_input_template.xlsx', sheet_name='지역간 연결')

print('지역간 연결 시트 상세 분석:')
print(f'총 행 수: {len(df)}')
print(f'총 열 수: {len(df.columns)}')
print()

print('첫 10행 전체 데이터:')
for i in range(min(10, len(df))):
    print(f'{i:2d}: {list(df.iloc[i])}')

print()
print('실제 연결 데이터 찾기:')
for i, row in df.iterrows():
    first_col = str(row.iloc[0])
    if '_' in first_col and len(first_col.split('_')) == 2:
        print(f'연결 데이터 발견 - 행 {i}: {list(row)}')
        if i > 10:  # 처음 몇 개만 출력
            break 