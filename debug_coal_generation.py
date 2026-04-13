import sys
sys.path.insert(0, 'src')
import openpyxl
from process_regional_excel import RegionalExcelProcessor

proc = RegionalExcelProcessor('interface.xlsx')
proc.output_path = 'integrated_input_data.xlsx'

wb = openpyxl.load_workbook('interface.xlsx', data_only=True)

print("=" * 80)
print("단계별 디버깅")
print("=" * 80)

# 1. Coal_DB 읽기
print("\n1. Coal_DB 읽기...")
success = proc.read_coal_db(wb)
print(f"   Result: {success}")
if proc.coal_db_generators is not None:
    print(f"   Coal_DB 발전기 수: {len(proc.coal_db_generators)}")

# 2. 선택된 지역 읽기
print("\n2. 선택된 지역 읽기...")
success = proc.read_selected_regions(wb)
print(f"   Result: {success}")
print(f"   선택된 지역: {proc.selected_regions}")

# 3. 지역별 데이터 읽기 (Coal 파라미터 저장 및 Coal_DB 발전기 추가)
print("\n3. 지역별 데이터 읽기...")
success = proc.read_regional_data(wb)
print(f"   Result: {success}")

# 4. Data manager 확인
print("\n4. Data manager에 저장된 발전기 확인...")
if hasattr(proc.data_manager, 'regions'):
    for region in proc.selected_regions:
        if region in proc.data_manager.regions:
            region_data = proc.data_manager.regions[region]
            if 'generators' in region_data:
                gens = region_data['generators']
                print(f"   {region}: {len(gens)}개 발전기")
                # 석탄 발전기 확인
                coal_count = sum(1 for g in gens if '당진' in g.get('name', '') or '영흥' in g.get('name', ''))
                if coal_count > 0:
                    print(f"      → 석탄발전기: {coal_count}개")
                    # 샘플 출력
                    for g in gens[:3]:
                        if '당진' in g.get('name', '') or '영흥' in g.get('name', ''):
                            print(f"         {g.get('name')}: {g.get('p_nom')} MW")

# 5. Merge data 확인
print("\n5. Merge data 실행...")
merged_data = proc.data_manager.merge_data()
if 'generators' in merged_data:
    print(f"   총 발전기 수: {len(merged_data['generators'])}")
    # 석탄 발전기 확인
    coal_gens = [g for idx, g in merged_data['generators'].iterrows() 
                 if '당진' in str(g.get('name', '')) or '영흥' in str(g.get('name', ''))]
    print(f"   석탄발전기 수: {len(coal_gens)}")
    if len(coal_gens) > 0:
        print("   샘플:")
        for i, g in enumerate(coal_gens[:5]):
            print(f"      {g['name']}: {g.get('p_nom')} MW")

wb.close()
