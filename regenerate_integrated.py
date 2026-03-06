import sys
sys.path.insert(0, 'src')
import openpyxl
from process_regional_excel import RegionalExcelProcessor
import os

# 기존 파일 삭제
if os.path.exists('integrated_input_data.xlsx'):
    print("기존 integrated_input_data.xlsx 삭제 중...")
    os.remove('integrated_input_data.xlsx')
    print("  삭제 완료")

proc = RegionalExcelProcessor('interface.xlsx')
proc.output_path = 'integrated_input_data.xlsx'

wb = openpyxl.load_workbook('interface.xlsx', data_only=True)

print("\n" + "=" * 80)
print("통합 데이터 재생성")
print("=" * 80)

# 1~5단계 실행
proc.read_coal_db(wb)
proc.read_selected_regions(wb)
proc.read_regional_data(wb)
proc.read_connections(wb)
proc.read_timeseries(wb)
proc.read_patterns(wb)

wb.close()

# 6. 통합 데이터 생성 및 저장
print("\n통합 데이터 생성 및 저장 중...")
success = proc.create_integrated_data()
print(f"Result: {success}")

if success and os.path.exists('integrated_input_data.xlsx'):
    print("\n파일 생성 확인...")
    import pandas as pd
    df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
    print(f"  총 발전기 수: {len(df_gens)}")
    
    coal_gens = df_gens[df_gens['name'].str.contains('당진|영흥|하동|삼척|태안', na=False)]
    print(f"  석탄발전기 수: {len(coal_gens)}")
    
    if len(coal_gens) > 0:
        print("  [SUCCESS] 석탄발전기가 정상적으로 포함되었습니다!")
    else:
        print("  [ERROR] 여전히 석탄발전기가 없습니다!")
