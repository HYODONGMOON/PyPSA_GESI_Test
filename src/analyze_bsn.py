import pandas as pd

# BSN 지역 데이터 구조 분석
df = pd.read_excel('regional_input_template.xlsx', sheet_name='지역_BSN')

print('BSN 지역 데이터 구조 분석:')
print(f'총 행 수: {len(df)}')
print(f'총 열 수: {len(df.columns)}')
print()

for i, row in df.iterrows():
    if i > 30:  # 처음 30행만 분석
        break
    
    first_col = str(row.iloc[0]).strip()
    second_col = str(row.iloc[1]).strip() if len(row) > 1 else ''
    
    if first_col and first_col != 'nan':
        print(f'{i:2d}: [{first_col[:15]}] [{second_col[:15]}] - {[str(cell)[:10] for cell in row[:5]]}') 