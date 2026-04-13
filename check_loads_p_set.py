import pandas as pd
import numpy as np

# loads 시트 읽기
df = pd.read_excel('integrated_input_data.xlsx', sheet_name='loads')

print('='*70)
print('Loads 시트 정보')
print('='*70)
print(f'행 수: {len(df)}')
print(f'\n컬럼: {list(df.columns)}')

print(f'\n=== p_set 통계 ===')
print(df['p_set'].describe())

print(f'\n=== 샘플 (처음 10행) ===')
print(df[['name', 'bus', 'p_set']].head(10).to_string(index=False))

# EL 부하만 필터링
el_loads = df[df['name'].str.contains('_Demand_EL', na=False)]
print(f'\n=== 전력(EL) 부하 분석 ===')
print(f'EL 부하 개수: {len(el_loads)}')
print(f'EL 부하 p_set 합계: {el_loads["p_set"].sum():.2f} MW')
print(f'EL 부하 p_set 평균: {el_loads["p_set"].mean():.2f} MW')

print(f'\n=== 전체 부하 p_set 합계 ===')
print(f'전체 부하 p_set 합계: {df["p_set"].sum():.2f} MW')

# 지역별 EL 부하
print(f'\n=== 지역별 EL 부하 ===')
for _, row in el_loads.iterrows():
    print(f'{row["name"]}: {row["p_set"]:.2f} MW')
