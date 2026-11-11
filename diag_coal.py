import sys, os, pandas as pd
sys.path.insert(0, '.')
sys.path.insert(0, 'src')

xls = pd.ExcelFile('integrated_input_data.xlsx')
gen_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
print('=== 로딩 전 generators ===')
print(f'총 {len(gen_df)}개, carriers: {dict(gen_df["carrier"].value_counts())}')
coal_before = gen_df[gen_df["name"].astype(str).str.endswith("_Coal")]
print(f'집계 석탄(*_Coal): {list(coal_before["name"])}')
print(f'집계 석탄 buses: {list(coal_before["bus"])}')
print(f'집계 석탄 carriers: {list(coal_before["carrier"])}')

from PyPSA_GUI import _replace_coal_generators_from_interface
input_data = {k: pd.read_excel('integrated_input_data.xlsx', sheet_name=k) for k in xls.sheet_names}
interface_path = os.path.abspath('interface.xlsx')

print('\n=== Coal_DB 대체 실행 ===')
input_data = _replace_coal_generators_from_interface(input_data, interface_path)

gen_after = input_data['generators']
coal_after = gen_after[gen_after['carrier'].astype(str).str.lower() == 'coal']
print(f'\n=== 대체 후 coal 발전기 ({len(coal_after)}개) ===')
if not coal_after.empty:
    print(coal_after[['name','bus','carrier','p_nom','efficiency','marginal_cost']].head(15).to_string())
else:
    print('석탄 발전기 없음!')

# 버스 목록도 확인
buses = input_data.get('buses', pd.DataFrame())
if not buses.empty and 'name' in buses.columns:
    print(f'\n실제 버스 이름 (EL 버스만): {[b for b in buses["name"].dropna() if str(b).endswith("_EL")]}')
