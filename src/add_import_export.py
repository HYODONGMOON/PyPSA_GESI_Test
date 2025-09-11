import pandas as pd
import os
import shutil
from datetime import datetime

def add_import_export():
    """
    네트워크에 수입/수출 노드를 추가하여 최적화 문제의 균형을 맞추는 함수
    
    1. 외부 연결용 버스 추가
    2. 각 지역에 수입/수출 링크 추가
    3. 적절한 비용 설정
    """
    print("수입/수출 노드 추가 중...")
    
    # 파일 경로
    input_file = 'integrated_input_data.xlsx'
    
    # 백업 파일 생성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = f'integrated_input_data_backup_iex_{timestamp}.xlsx'
    
    if not os.path.exists(input_file):
        print(f"오류: {input_file} 파일을 찾을 수 없습니다.")
        return False
    
    # 백업 생성
    shutil.copy2(input_file, backup_file)
    print(f"원본 파일을 {backup_file}로 백업했습니다.")
    
    try:
        # 데이터 로드
        with pd.ExcelFile(input_file) as xls:
            buses = pd.read_excel(xls, sheet_name='buses')
            links = pd.read_excel(xls, sheet_name='links')
            generators = pd.read_excel(xls, sheet_name='generators')
            loads = pd.read_excel(xls, sheet_name='loads')
            lines = pd.read_excel(xls, sheet_name='lines')
            stores = pd.read_excel(xls, sheet_name='stores')
        
        # 버스와 지역 코드 확인
        region_codes = []
        el_buses = []
        
        for bus in buses['name']:
            if '_EL' in bus:
                el_buses.append(bus)
                region_code = bus.split('_')[0]
                if region_code not in region_codes:
                    region_codes.append(region_code)
        
        print(f"발견된 지역 코드: {', '.join(region_codes)}")
        print(f"전력 버스 수: {len(el_buses)}개")
        
        # 외부 연결 버스 추가
        external_bus = {
            'name': 'External_Grid',
            'v_nom': 345,
            'carrier': 'AC',
            'x': 0,  # 임의의 좌표
            'y': 0   # 임의의 좌표
        }
        
        if 'External_Grid' not in buses['name'].values:
            buses = pd.concat([buses, pd.DataFrame([external_bus])], ignore_index=True)
            print("'External_Grid' 버스가 추가되었습니다.")
        
        # 각 지역에 수입/수출 링크 추가
        new_links = []
        added_count = 0
        
        # 수입 링크 설정값
        import_cost = 100    # 수입 비용 (높게 설정하여 내부 생산 우선)
        export_cost = -10    # 수출 수익 (낮게 설정하여 내부 소비 우선)
        link_capacity = 1000  # 링크 용량 (MW)
        
        for bus in el_buses:
            region_code = bus.split('_')[0]
            
            # 수입 링크 (External_Grid -> 지역)
            import_link_name = f"Import_{region_code}"
            if import_link_name not in links['name'].values:
                import_link = {
                    'name': import_link_name,
                    'bus0': 'External_Grid',
                    'bus1': bus,
                    'p_nom': link_capacity,
                    'p_nom_extendable': True,
                    'p_nom_min': 0,
                    'p_nom_max': 5000,  # 최대 수입 용량
                    'efficiency': 1.0,
                    'marginal_cost': import_cost,
                    'capital_cost': 500  # 설비 투자비용
                }
                new_links.append(import_link)
                added_count += 1
            
            # 수출 링크 (지역 -> External_Grid)
            export_link_name = f"Export_{region_code}"
            if export_link_name not in links['name'].values:
                export_link = {
                    'name': export_link_name,
                    'bus0': bus,
                    'bus1': 'External_Grid',
                    'p_nom': link_capacity,
                    'p_nom_extendable': True,
                    'p_nom_min': 0,
                    'p_nom_max': 5000,  # 최대 수출 용량
                    'efficiency': 0.95,  # 송전 손실 반영
                    'marginal_cost': export_cost,  # 수출은 음의 비용(=수익)으로 설정
                    'capital_cost': 500  # 설비 투자비용
                }
                new_links.append(export_link)
                added_count += 1
        
        # 새 링크 추가
        if new_links:
            links = pd.concat([links, pd.DataFrame(new_links)], ignore_index=True)
            print(f"{added_count}개의 수입/수출 링크가 추가되었습니다.")
        
        # 엑셀 파일에 저장
        with pd.ExcelWriter(input_file) as writer:
            buses.to_excel(writer, sheet_name='buses', index=False)
            links.to_excel(writer, sheet_name='links', index=False)
            generators.to_excel(writer, sheet_name='generators', index=False)
            loads.to_excel(writer, sheet_name='loads', index=False)
            lines.to_excel(writer, sheet_name='lines', index=False)
            stores.to_excel(writer, sheet_name='stores', index=False)
        
        print(f"\n{input_file} 파일이 성공적으로 수정되었습니다.")
        print("이제 PyPSA_GUI.py를 다시 실행해보세요.")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        
        # 백업에서 복원
        print(f"백업에서 복원 중...")
        shutil.copy2(backup_file, input_file)
        print(f"원본 파일이 백업에서 복원되었습니다.")
        
        return False

if __name__ == "__main__":
    add_import_export() 