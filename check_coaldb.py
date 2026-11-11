import openpyxl, sys
sys.path.insert(0, 'src')
from process_regional_excel import RegionalExcelProcessor

wb = openpyxl.load_workbook('interface.xlsx', data_only=True)
ws = wb['Coal_DB']

# D2 셀에서 선택연도 확인
print(f"Coal_DB D2 (선택연도): {ws['D2'].value}")

# 헤더 행 (row 5) 확인
print("\nRow 5 (헤더):", [ws.cell(5, c).value for c in range(1, 40)])

# 2030년 컬럼 찾기
year_row = [ws.cell(5, c).value for c in range(1, 50)]
print("\n연도 목록 (row 5):", [(i+1, v) for i, v in enumerate(year_row) if v and str(v).isdigit()])

# RegionalExcelProcessor로 읽기
proc = RegionalExcelProcessor(wb)
proc.read_coal_db()

print(f"\n=== Coal_DB 발전기 ({len(proc.coal_db_generators)}개) ===")
if not proc.coal_db_generators.empty:
    print(proc.coal_db_generators[['region','plant_name','unit','capacity_MW','efficiency']].to_string())
else:
    print("발전기 없음!")
    
# 지역별 파라미터
proc.read_regional_data(wb)
print(f"\n=== 지역별 coal 파라미터 ({len(proc.region_coal_params)}개 지역) ===")
for region, params in proc.region_coal_params.items():
    print(f"  {region}: bus={params.get('bus')}, mc={params.get('marginal_cost')}, eff={params.get('efficiency')}")
