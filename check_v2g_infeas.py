import pandas as pd

# V2G 링크 p_min_pu 확인
links = pd.read_excel('integrated_input_data.xlsx', sheet_name='links')
v2g = links[links['name'].str.contains('V2G', na=False)]
cols = [c for c in v2g.columns if any(k in c for k in ['p_nom','p_min','p_max','marginal'])]
print('=== V2G link bounds ===')
print(v2g[['name'] + cols].to_string())

# timeseries 시트에 V2G p_max_pu, p_min_pu 있는지 확인
ts = pd.read_excel('integrated_input_data.xlsx', sheet_name='timeseries')
print('\n=== timeseries columns (V2G 관련) ===')
v2g_ts_cols = [c for c in ts.columns if 'V2G' in str(c) or 'v2g' in str(c).lower()]
print(v2g_ts_cols if v2g_ts_cols else 'V2G 관련 timeseries 없음')

# stores e_min_pu, e_max_pu 확인
stores = pd.read_excel('integrated_input_data.xlsx', sheet_name='stores')
ev_s = stores[stores['name'].str.contains('EV', na=False)]
scols = [c for c in ev_s.columns if any(k in c for k in ['e_nom','e_min','e_max','e_cyclic','e_initial'])]
print('\n=== EV store bounds ===')
print(ev_s[['name'] + scols].to_string())
