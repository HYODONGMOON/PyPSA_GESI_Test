import pandas as pd

df = pd.read_excel('integrated_input_data.xlsx', sheet_name='buses')
print("=== 버스 시트 컬럼:", list(df.columns))
print("=== 버스 이름 (EL 관련):")
bus_names = df['name'].dropna().astype(str).tolist() if 'name' in df.columns else []
for b in bus_names:
    if 'EL' in b or 'LNG' in b:
        print(f"  {b}")

print("\n=== 전체 버스 목록:")
for b in bus_names:
    print(f"  {b}")
