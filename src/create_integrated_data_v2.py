import pandas as pd
import numpy as np
import os
from datetime import datetime

def get_selected_regions():
    """지역 선택 시트에서 선택된 지역 목록 가져오기"""
    try:
        df = pd.read_excel('regional_input_template.xlsx', sheet_name='지역 선택')
        
        # 헤더 찾기
        header_row = None
        for i, row in df.iterrows():
            if '코드' in str(row.iloc[0]) and '지역명' in str(row.iloc[1]):
                header_row = i
                break
        
        if header_row is None:
            print("헤더를 찾을 수 없습니다.")
            return []
        
        # 헤더 이후 데이터 읽기
        data_df = df.iloc[header_row+1:].copy()
        data_df.columns = ['코드', '지역명', '영문명', '선택', '인구', '면적']
        
        # 선택된 지역 필터링
        selected_regions = []
        for _, row in data_df.iterrows():
            if pd.notna(row['코드']) and str(row['선택']).strip().upper() == 'O':
                selected_regions.append(str(row['코드']).strip())
        
        print(f"선택된 지역: {selected_regions}")
        return selected_regions
        
    except Exception as e:
        print(f"지역 선택 읽기 오류: {str(e)}")
        return []

def parse_regional_data(region_code):
    """특정 지역의 데이터를 파싱하여 구조화된 데이터 반환"""
    try:
        sheet_name = f'지역_{region_code}'
        df = pd.read_excel('regional_input_template.xlsx', sheet_name=sheet_name)
        
        regional_data = {
            'buses': [],
            'generators': [],
            'loads': [],
            'stores': [],
            'links': []
        }
        
        current_section = None
        header_found = False
        
        for i, row in df.iterrows():
            row_data = [str(cell).strip() if pd.notna(cell) else '' for cell in row]
            first_col = row_data[0]
            
            # 섹션 식별
            if '버스' in first_col and '이름' in str(row.iloc[1]):
                current_section = 'buses'
                header_found = True
                continue
            elif '발전기' in first_col and '이름' in str(row.iloc[1]):
                current_section = 'generators'
                header_found = True
                continue
            elif '부하' in first_col and '이름' in str(row.iloc[1]):
                current_section = 'loads'
                header_found = True
                continue
            elif '저장장치' in first_col and '이름' in str(row.iloc[1]):
                current_section = 'stores'
                header_found = True
                continue
            elif '링크' in first_col and '이름' in str(row.iloc[1]):
                current_section = 'links'
                header_found = True
                continue
            
            # 데이터 행 처리
            if current_section and header_found and first_col and first_col not in ['', 'NaN', 'nan']:
                # 실제 데이터인지 확인
                if not any(keyword in first_col for keyword in ['설명', '입력', '예시', '주의', '섹션']):
                    if current_section == 'buses':
                        # 버스 데이터: 이름, 전압, 캐리어, X좌표, Y좌표
                        if len(row_data) >= 3:
                            bus_data = {
                                'name': f"{region_code}_{first_col}",
                                'v_nom': row_data[2] if row_data[2] and row_data[2] != '' else 345,
                                'carrier': row_data[3] if len(row_data) > 3 and row_data[3] else 'AC',
                                'x': row_data[4] if len(row_data) > 4 and row_data[4] else 0,
                                'y': row_data[5] if len(row_data) > 5 and row_data[5] else 0
                            }
                            regional_data['buses'].append(bus_data)
                    
                    elif current_section == 'generators':
                        # 발전기 데이터
                        if len(row_data) >= 4:
                            gen_data = {
                                'name': f"{region_code}_{first_col}",
                                'bus': f"{region_code}_{row_data[1]}" if row_data[1] else f"{region_code}_EL",
                                'carrier': row_data[2] if row_data[2] else 'electricity',
                                'p_nom': float(row_data[3]) if row_data[3] and str(row_data[3]).replace('.','').isdigit() else 100,
                                'p_nom_extendable': True,
                                'marginal_cost': float(row_data[7]) if len(row_data) > 7 and row_data[7] and str(row_data[7]).replace('.','').isdigit() else 50,
                                'capital_cost': float(row_data[8]) if len(row_data) > 8 and row_data[8] and str(row_data[8]).replace('.','').isdigit() else 1000000,
                                'efficiency': float(row_data[9]) if len(row_data) > 9 and row_data[9] and str(row_data[9]).replace('.','').isdigit() else 0.4
                            }
                            regional_data['generators'].append(gen_data)
                    
                    elif current_section == 'loads':
                        # 부하 데이터
                        if len(row_data) >= 3:
                            load_data = {
                                'name': f"{region_code}_{first_col}",
                                'bus': f"{region_code}_{row_data[1]}" if row_data[1] else f"{region_code}_EL",
                                'p_set': float(row_data[2]) if row_data[2] and str(row_data[2]).replace('.','').isdigit() else 1000
                            }
                            regional_data['loads'].append(load_data)
                    
                    elif current_section == 'stores':
                        # 저장장치 데이터
                        if len(row_data) >= 3:
                            store_data = {
                                'name': f"{region_code}_{first_col}",
                                'bus': f"{region_code}_{row_data[1]}" if row_data[1] else f"{region_code}_EL",
                                'carrier': row_data[2] if row_data[2] else 'electricity',
                                'e_nom': float(row_data[3]) if len(row_data) > 3 and row_data[3] and str(row_data[3]).replace('.','').isdigit() else 1000,
                                'e_nom_extendable': True,
                                'efficiency_store': 0.9,
                                'efficiency_dispatch': 0.9
                            }
                            regional_data['stores'].append(store_data)
                    
                    elif current_section == 'links':
                        # 링크 데이터
                        if len(row_data) >= 4:
                            link_data = {
                                'name': f"{region_code}_{first_col}",
                                'bus0': f"{region_code}_{row_data[1]}" if row_data[1] else f"{region_code}_EL",
                                'bus1': f"{region_code}_{row_data[2]}" if row_data[2] else f"{region_code}_H2",
                                'efficiency': float(row_data[3]) if row_data[3] and str(row_data[3]).replace('.','').isdigit() else 0.7,
                                'p_nom': float(row_data[4]) if len(row_data) > 4 and row_data[4] and str(row_data[4]).replace('.','').isdigit() else 100,
                                'p_nom_extendable': True
                            }
                            regional_data['links'].append(link_data)
        
        return regional_data
        
    except Exception as e:
        print(f"지역 {region_code} 데이터 파싱 오류: {str(e)}")
        return {}

def create_integrated_data():
    """선택된 지역들의 데이터를 통합하여 integrated_input_data.xlsx 생성"""
    
    try:
        # 선택된 지역 가져오기
        selected_regions = get_selected_regions()
        if not selected_regions:
            print("선택된 지역이 없습니다.")
            return False
        
        # 통합 데이터 구조 초기화
        all_buses = []
        all_generators = []
        all_loads = []
        all_stores = []
        all_links = []
        
        # 각 지역 데이터 수집
        for region in selected_regions:
            print(f"\n{region} 지역 데이터 처리 중...")
            regional_data = parse_regional_data(region)
            
            if regional_data:
                all_buses.extend(regional_data.get('buses', []))
                all_generators.extend(regional_data.get('generators', []))
                all_loads.extend(regional_data.get('loads', []))
                all_stores.extend(regional_data.get('stores', []))
                all_links.extend(regional_data.get('links', []))
                
                print(f"  buses: {len(regional_data.get('buses', []))}개")
                print(f"  generators: {len(regional_data.get('generators', []))}개")
                print(f"  loads: {len(regional_data.get('loads', []))}개")
                print(f"  stores: {len(regional_data.get('stores', []))}개")
                print(f"  links: {len(regional_data.get('links', []))}개")
        
        # DataFrame 생성
        buses_df = pd.DataFrame(all_buses)
        generators_df = pd.DataFrame(all_generators)
        loads_df = pd.DataFrame(all_loads)
        stores_df = pd.DataFrame(all_stores)
        links_df = pd.DataFrame(all_links)
        
        # 기본 시트들 생성
        timeseries_df = pd.DataFrame({
            'start_time': ['2024-01-01 00:00:00'],
            'end_time': ['2024-12-31 23:00:00'],
            'frequency': ['h']
        })
        
        # 빈 lines 시트
        lines_df = pd.DataFrame()
        
        # integrated_input_data.xlsx 생성
        output_file = "integrated_input_data.xlsx"
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            buses_df.to_excel(writer, sheet_name='buses', index=False)
            generators_df.to_excel(writer, sheet_name='generators', index=False)
            loads_df.to_excel(writer, sheet_name='loads', index=False)
            stores_df.to_excel(writer, sheet_name='stores', index=False)
            links_df.to_excel(writer, sheet_name='links', index=False)
            lines_df.to_excel(writer, sheet_name='lines', index=False)
            timeseries_df.to_excel(writer, sheet_name='timeseries', index=False)
        
        print(f"\n=== 통합 결과 ===")
        print(f"총 buses: {len(buses_df)}개")
        print(f"총 generators: {len(generators_df)}개")
        print(f"총 loads: {len(loads_df)}개")
        print(f"총 stores: {len(stores_df)}개")
        print(f"총 links: {len(links_df)}개")
        
        print(f"\n{output_file} 파일이 성공적으로 생성되었습니다.")
        
        # 파일 수정 시간 확인
        mod_time = os.path.getmtime(output_file)
        mod_datetime = datetime.fromtimestamp(mod_time)
        print(f"파일 수정 시간: {mod_datetime}")
        
        return True
        
    except Exception as e:
        print(f"통합 데이터 생성 중 오류: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=== 지역별 데이터 통합하여 integrated_input_data.xlsx 생성 (v2) ===")
    create_integrated_data() 