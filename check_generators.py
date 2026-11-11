import pandas as pd

gen_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
print("=== 발전기 전체 목록 ===")
print(gen_df[['name','bus','carrier','p_nom','efficiency','marginal_cost']].to_string())

print("\n=== electricity carrier 발전기 요약 ===")
elec = gen_df[gen_df['carrier']=='electricity']
for _, row in elec.iterrows():
    print(f"  {row['name']}: bus={row['bus']}, p_nom={row['p_nom']}, mc={row['marginal_cost']}, eff={row['efficiency']}")
