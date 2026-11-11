import openpyxl, sys
sys.path.insert(0, 'src')
wb = openpyxl.load_workbook('interface.xlsx', data_only=True)

print("=== 지역별 발전원 한계비용 ===")
for sname in wb.sheetnames:
    if '_' in sname:
        region = sname.split('_')[-1]
        if len(region) == 3 and region.isalpha():
            ws = wb[sname]
            costs = {}
            for row in ws.iter_rows(min_row=1, max_row=60, values_only=True):
                name = row[0]
                if name in ('Coal', 'LNG', 'Nuclear', 'nuclear', '원자력', 'Gas'):
                    mc = row[7]
                    carrier = row[2]
                    costs[str(name)] = (mc, carrier)
            if costs:
                print(f"  {region}: " + ", ".join(f"{k}={v[0]}(carrier={v[1]})" for k, v in costs.items()))

print("\n=== 결과 파일에서 coal 발전기 발전량 확인 ===")
import os, glob
result_files = glob.glob('results_*.xlsx')
if result_files:
    latest = max(result_files, key=os.path.getmtime)
    print(f"최신 결과 파일: {latest}")
    import pandas as pd
    xls = pd.ExcelFile(latest)
    sheets = xls.sheet_names
    print(f"시트 목록: {[s for s in sheets if 'Coal' in s or 'Generator' in s]}")
    if 'Generator_Output_Coal' in sheets:
        coal_df = pd.read_excel(latest, sheet_name='Generator_Output_Coal')
        print("Generator_Output_Coal (일부):")
        if not coal_df.empty:
            total_gen = coal_df.select_dtypes(include='number').sum().sum()
            print(f"  총 석탄 발전량: {total_gen:.1f} MWh")
            print(coal_df.sum(numeric_only=True).sort_values(ascending=False).head(5).to_string())
