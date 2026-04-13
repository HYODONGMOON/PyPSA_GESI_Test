import pandas as pd

df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

print('='*70)
print('Generator Carrier 최종 확인')
print('='*70)

print('\nCarrier별 집계:')
print(df.groupby('carrier')['p_nom'].agg(['count', 'sum']).to_string())

print('\n\n샘플 (처음 25개):')
print(df[['name', 'bus', 'carrier', 'p_nom', 'marginal_cost']].head(25).to_string(index=False))

print('\n\n문제가 되었던 발전기 확인:')
problem_gens = ['BSN_Coal', 'BSN_Hydro', 'CND_Coal', 'CBD_Hydro', 'BSN_LNG', 'BSN_Nuclear']
for g in problem_gens:
    row = df[df['name']==g]
    if not row.empty:
        print(f'  {g}: carrier={row.iloc[0]["carrier"]}, p_nom={row.iloc[0]["p_nom"]}')
    else:
        print(f'  {g}: NOT FOUND')

print('\n' + '='*70)
