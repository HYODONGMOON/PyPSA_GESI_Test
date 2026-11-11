import pandas as pd

# 1. Stores 경계 확인
df_s = pd.read_excel('integrated_input_data.xlsx', sheet_name='stores')
print('=== Stores 경계 이상 ===')
found = False
for _, row in df_s.iterrows():
    e = float(row.get('e_nom', 0) or 0)
    mx = row.get('e_nom_max', None)
    mn = row.get('e_nom_min', None)
    ext = bool(row.get('e_nom_extendable', False))
    if mx is not None and not pd.isna(mx):
        mx_f = float(mx)
        if mx_f == 0 and ext:
            print(f'  [주의] {row["name"]}: e_nom_max=0 + extendable=True (완전 차단)')
            found = True
        if e > mx_f + 0.01:
            print(f'  [오류] {row["name"]}: e_nom({e:.0f}) > e_nom_max({mx_f:.0f}) -- 하한>상한!')
            found = True
    if mn is not None and not pd.isna(mn) and mx is not None and not pd.isna(mx):
        if float(mn) > float(mx) + 0.01:
            print(f'  [오류] {row["name"]}: e_nom_min({float(mn):.0f}) > e_nom_max({float(mx):.0f})')
            found = True
if not found:
    print('  이상 없음')

# 2. Generators 경계 확인
df_g = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
print('\n=== Generators p_nom vs p_nom_max ===')
found2 = False
for _, row in df_g.iterrows():
    p = float(row.get('p_nom', 0) or 0)
    mx = row.get('p_nom_max', None)
    mn = row.get('p_nom_min', None)
    ext = bool(row.get('p_nom_extendable', False))
    if ext and mx is not None and not pd.isna(mx):
        if mn is not None and not pd.isna(mn) and float(mn) > float(mx) + 0.01:
            print(f'  [오류] {row["name"]}: p_nom_min({float(mn):.0f}) > p_nom_max({float(mx):.0f})')
            found2 = True
    if not ext and mx is not None and not pd.isna(mx):
        if p > float(mx) + 0.01:
            print(f'  [오류] {row["name"]}: p_nom({p:.0f}) > p_nom_max({float(mx):.0f})')
            found2 = True
if not found2:
    print('  이상 없음')

# 3. 설비 용량 합계
print('\n=== 발전원별 용량 합계 ===')
elec = df_g[df_g['carrier'] == 'electricity']
coal = df_g[df_g['carrier'] == 'coal']
print(f'LNG:     {elec[elec["name"].str.contains("LNG",na=False)]["p_nom"].sum():.0f} MW')
print(f'Nuclear: {elec[elec["name"].str.contains("Nuclear",na=False)]["p_nom"].sum():.0f} MW')
print(f'Coal:    {coal["p_nom"].sum():.0f} MW')
print(f'PV:      {df_g[df_g["carrier"]=="solar"]["p_nom"].sum():.0f} MW')
print(f'WT:      {df_g[df_g["carrier"]=="wind"]["p_nom"].sum():.0f} MW')
print(f'합계 공급 최대: {elec["p_nom"].sum() + coal["p_nom"].sum() + df_g[df_g["carrier"].isin(["solar","wind"])]["p_nom"].sum():.0f} MW')

# 4. 전력 수요 확인
try:
    df_load = pd.read_excel('integrated_input_data.xlsx', sheet_name='loads')
    el_load = df_load[df_load.get('carrier', pd.Series(dtype=str)).eq('electricity') if 'carrier' in df_load.columns else pd.Series(True, index=df_load.index)]
    print(f'\n전력 수요 (p_set 합계): {df_load["p_set"].sum():.0f} MW')
    print(df_load[['name','p_set']].head(10).to_string())
except Exception as ex:
    print(f'loads 시트 오류: {ex}')

# 5. Pumped hydro 확인
print('\n=== Pumped Hydro 상태 ===')
pumped = df_s[df_s['name'].str.contains('Pumped', case=False, na=False)]
cols = [c for c in ['name','e_nom','e_nom_extendable','e_nom_max','e_nom_min'] if c in df_s.columns]
print(pumped[cols].to_string())
