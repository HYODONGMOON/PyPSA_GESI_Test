import pandas as pd

xls = pd.ExcelFile('results/rolling_horizon_20260204_132341/rolling_horizon_result_20260204_132341.xlsx')
lines_df = pd.read_excel(xls, 'Line_Flow')

print(f'Line_Flow sheet rows: {len(lines_df)}')
print(f'Unique lines: {lines_df["Line"].nunique()}')
print('\nLine_Flow sample (first 10):')
print(lines_df.head(10))
print('\nUnique line names sample:')
print(lines_df['Line'].unique()[:30])
