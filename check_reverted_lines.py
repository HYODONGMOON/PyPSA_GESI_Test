import pandas as pd
import openpyxl

# interface.xlsx 파일 확인
wb = openpyxl.load_workbook('interface.xlsx', read_only=True)
print("Sheet names:", wb.sheetnames)
wb.close()

# 지역간 연결 시트 읽기
try:
    df = pd.read_excel('interface.xlsx', sheet_name='지역간 연결')
    print(f"\n지역간 연결 시트:")
    print(f"  Lines count: {len(df)}")
    print(f"  Columns: {df.columns.tolist()}")
    print(f"\n상위 10개:")
    print(df.head(10))
except Exception as e:
    print(f"Error reading 지역간 연결: {e}")
