import pandas as pd

# 지역간 연결 시트 분석
df = pd.read_excel('regional_input_template.xlsx', sheet_name='지역간 연결')

print('지역간 연결 시트 분석:')
print(f'총 행 수: {len(df)}')
print(f'총 열 수: {len(df.columns)}')
print(f'컬럼명: {list(df.columns)}')
print()

print('첫 15행 데이터:')
for i in range(min(15, len(df))):
    row_data = [str(cell)[:15] for cell in df.iloc[i][:8]]
    print(f'{i:2d}: {row_data}')

print()
print('실제 데이터 부분 (헤더 제외):')
# 헤더 찾기
header_row = None
for i, row in df.iterrows():
    if '이름' in str(row.iloc[0]) or 'name' in str(row.iloc[0]).lower():
        header_row = i
        break

if header_row is not None:
    print(f'헤더 행: {header_row}')
    print('헤더:', [str(cell) for cell in df.iloc[header_row][:8]])
    print('데이터 샘플:')
    for i in range(header_row+1, min(header_row+6, len(df))):
        if pd.notna(df.iloc[i, 0]):
            print(f'{i}: {[str(cell) for cell in df.iloc[i][:8]]}')
else:
    print('헤더를 찾을 수 없습니다.') 