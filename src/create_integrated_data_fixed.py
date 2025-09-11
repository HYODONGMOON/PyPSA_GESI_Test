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

def parse_inter_region_connections():
    """지역간 연결 시트에서 lines 데이터 파싱 (수정된 버전)"""
    try:
        df = pd.read_excel('regional_input_template.xlsx', sheet_name='지역간 연결')
        
        lines_data = []
        
        # 실제 연결 데이터 찾기 (첫 번째 컬럼에 '_'가 포함된 행)
        for i, row in df.iterrows():
            first_col = str(row.iloc[0])
            if pd.notna(row.iloc[0]) and '_' in first_col and len(first_col.split('_')) == 2:
                try:
                    # 연결 이름에서 지역 코드 추출
                    regions = first_col.split('_')
                    bus0_name = f"{regions[0]}_EL"
                    bus1_name = f"{regions[1]}_EL"
                    
                    # 나머지 데이터 파싱
                    line_data = {
                        'name': first_col,
                        'bus0': bus0_name,
                        'bus1': bus1_name,
                        'carrier': 'AC',
                        'x': 0.01,  # 기본값
                        'r': 0.001,  # 기본값
                        's_nom': 1000,  # 기본값
                        'length': 100,  # 기본값
                        'v_nom': 345
                    }
                    
                    # 실제 데이터가 있으면 사용
                    if len(row) > 4 and pd.notna(row.iloc[4]):
                        try:
                            line_data['length'] = float(row.iloc[4])
                        except:
                            pass
                    
                    if len(row) > 5 and pd.notna(row.iloc[5]):
                        try:
                            line_data['s_nom'] = float(row.iloc[5])
                        except:
                            pass
                    
                    if len(row) > 6 and pd.notna(row.iloc[6]):
                        try:
                            line_data['x'] = float(row.iloc[6])
                        except:
                            pass
                    
                    if len(row) > 7 and pd.notna(row.iloc[7]):
                        try:
                            line_data['r'] = float(row.iloc[7])
                        except:
                            pass
                    
                    lines_data.append(line_data)
                    print(f"지역간 연결 추가: {first_col} ({bus0_name} -> {bus1_name})")
                    
                except Exception as e:
                    print(f"연결 {first_col} 파싱 중 오류: {str(e)}")
                    continue
        
        print(f"지역간 연결 lines: {len(lines_data)}개 파싱됨")
        return lines_data
        
    except Exception as e:
        print(f"지역간 연결 파싱 오류: {str(e)}")
        return []

def parse_constraints():
    """constraints 시트에서 제약조건 데이터 파싱"""
    try:
        df = pd.read_excel('regional_input_template.xlsx', sheet_name='constraints')
        
        constraints_data = []
        
        # 헤더 찾기
        header_row = None
        for i, row in df.iterrows():
            if '이름' in str(row.iloc[0]) or 'name' in str(row.iloc[0]).lower():
                header_row = i
                break
        
        if header_row is not None:
            # 헤더 이후 데이터 처리
            for i in range(header_row+1, len(df)):
                row = df.iloc[i]
                if pd.notna(row.iloc[0]) and str(row.iloc[0]).strip():
                    constraint_data = {
                        'name': str(row.iloc[0]),
                        'type': str(row.iloc[1]) if pd.notna(row.iloc[1]) else 'GlobalConstraint',
                        'carrier_attribute': str(row.iloc[2]) if pd.notna(row.iloc[2]) else 'co2_emissions',
                        'sense': str(row.iloc[3]) if pd.notna(row.iloc[3]) else '<=',
                        'constant': float(row.iloc[4]) if pd.notna(row.iloc[4]) else 1000000
                    }
                    constraints_data.append(constraint_data)
        
        # 기본 CO2 제약조건이 없으면 추가
        if not constraints_data:
            constraints_data.append({
                'name': 'CO2Limit',
                'type': 'GlobalConstraint',
                'carrier_attribute': 'co2_emissions',
                'sense': '<=',
                'constant': 1000000
            })
        
        print(f"Constraints: {len(constraints_data)}개 파싱됨")
        return constraints_data
        
    except Exception as e:
        print(f"Constraints 파싱 오류: {str(e)}")
        # 기본 제약조건 반환
        return [{
            'name': 'CO2Limit',
            'type': 'GlobalConstraint',
            'carrier_attribute': 'co2_emissions',
            'sense': '<=',
            'constant': 1000000
        }]

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
        
        for i, row in df.iterrows():
            first_col = str(row.iloc[0]).strip()
            second_col = str(row.iloc[1]).strip() if len(row) > 1 else ''
            
            # 섹션 헤더 식별
            if first_col == '버스':
                current_section = 'buses'
                continue
            elif first_col == '발전기':
                current_section = 'generators'
                continue
            elif first_col == '부하':
                current_section = 'loads'
                continue
            elif first_col == '저장장치':
                current_section = 'stores'
                continue
            elif first_col == '링크':
                current_section = 'links'
                continue
            
            # 헤더 행 건너뛰기
            if first_col == '이름':
                continue
            
            # 데이터 행 처리
            if current_section and first_col and first_col not in ['nan', '', 'NaN']:
                if current_section == 'buses':
                    bus_data = {
                        'name': f"{region_code}_{first_col}",
                        'v_nom': float(second_col) if second_col and second_col != 'nan' else 345,
                        'carrier': str(row.iloc[2]) if len(row) > 2 and pd.notna(row.iloc[2]) else 'AC',
                        'x': float(row.iloc[3]) if len(row) > 3 and pd.notna(row.iloc[3]) else 0,
                        'y': float(row.iloc[4]) if len(row) > 4 and pd.notna(row.iloc[4]) else 0
                    }
                    regional_data['buses'].append(bus_data)
                
                elif current_section == 'generators':
                    try:
                        p_nom_val = float(row.iloc[3]) if len(row) > 3 and pd.notna(row.iloc[3]) else 100
                        if p_nom_val > 0:
                            gen_data = {
                                'name': f"{region_code}_{first_col}",
                                'bus': f"{region_code}_{second_col}" if second_col and second_col != 'nan' else f"{region_code}_EL",
                                'carrier': str(row.iloc[2]) if len(row) > 2 and pd.notna(row.iloc[2]) else 'electricity',
                                'p_nom': p_nom_val,
                                'p_nom_extendable': str(row.iloc[4]).lower() == 'true' if len(row) > 4 and pd.notna(row.iloc[4]) else False,
                                'marginal_cost': float(row.iloc[7]) if len(row) > 7 and pd.notna(row.iloc[7]) else 50,
                                'capital_cost': float(row.iloc[8]) if len(row) > 8 and pd.notna(row.iloc[8]) else 1000000,
                                'efficiency': float(row.iloc[9]) if len(row) > 9 and pd.notna(row.iloc[9]) else 0.4
                            }
                            regional_data['generators'].append(gen_data)
                    except:
                        pass
                
                elif current_section == 'loads':
                    try:
                        load_data = {
                            'name': f"{region_code}_{first_col}",
                            'bus': f"{region_code}_{second_col}" if second_col and second_col != 'nan' else f"{region_code}_EL",
                            'p_set': float(row.iloc[3]) if len(row) > 3 and pd.notna(row.iloc[3]) else 1000
                        }
                        regional_data['loads'].append(load_data)
                    except:
                        pass
                
                elif current_section == 'stores':
                    try:
                        store_data = {
                            'name': f"{region_code}_{first_col}",
                            'bus': f"{region_code}_{second_col}" if second_col and second_col != 'nan' else f"{region_code}_EL",
                            'carrier': str(row.iloc[2]) if len(row) > 2 and pd.notna(row.iloc[2]) else 'electricity',
                            'e_nom': float(row.iloc[3]) if len(row) > 3 and pd.notna(row.iloc[3]) else 1000,
                            'e_nom_extendable': str(row.iloc[4]).lower() == 'true' if len(row) > 4 and pd.notna(row.iloc[4]) else True,
                            'e_cyclic': True,
                            'standing_loss': 0.0,
                            'efficiency_store': 0.9,
                            'efficiency_dispatch': 0.9,
                            'e_initial': 0,
                            'e_nom_max': float(row.iloc[5]) if len(row) > 5 and pd.notna(row.iloc[5]) else 100000
                        }
                        regional_data['stores'].append(store_data)
                    except:
                        pass
                
                elif current_section == 'links':
                    try:
                        link_data = {
                            'name': f"{region_code}_{first_col}",
                            'bus0': f"{region_code}_{second_col}" if second_col and second_col != 'nan' else f"{region_code}_EL",
                            'bus1': f"{region_code}_{row.iloc[2]}" if len(row) > 2 and pd.notna(row.iloc[2]) else f"{region_code}_H2",
                            'efficiency': float(row.iloc[3]) if len(row) > 3 and pd.notna(row.iloc[3]) else 0.7,
                            'p_nom': float(row.iloc[4]) if len(row) > 4 and pd.notna(row.iloc[4]) else 100,
                            'p_nom_extendable': True
                        }
                        regional_data['links'].append(link_data)
                    except:
                        pass
        
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
        
        # 지역간 연결 데이터 가져오기
        print(f"\n지역간 연결 데이터 처리 중...")
        lines_data = parse_inter_region_connections()
        
        # 제약조건 데이터 가져오기
        print(f"\n제약조건 데이터 처리 중...")
        constraints_data = parse_constraints()
        
        # DataFrame 생성
        buses_df = pd.DataFrame(all_buses)
        generators_df = pd.DataFrame(all_generators)
        loads_df = pd.DataFrame(all_loads)
        stores_df = pd.DataFrame(all_stores)
        links_df = pd.DataFrame(all_links)
        lines_df = pd.DataFrame(lines_data)
        constraints_df = pd.DataFrame(constraints_data)
        
        # 기본 시트들 생성
        timeseries_df = pd.DataFrame({
            'start_time': ['2024-01-01 00:00:00'],
            'end_time': ['2024-12-31 23:00:00'],
            'frequency': ['h']
        })
        
        # 재생에너지 패턴 (기존 데이터 사용)
        try:
            renewable_patterns_df = pd.read_excel('regional_input_template.xlsx', sheet_name='renewable_patterns')
        except:
            # 기본 패턴 생성
            renewable_patterns_df = pd.DataFrame({
                'hour': range(1, 8761),
                'PV': np.random.uniform(0, 1, 8760),
                'WT': np.random.uniform(0, 1, 8760)
            })
        
        # 부하 패턴 (기존 데이터 사용)
        try:
            load_patterns_df = pd.read_excel('regional_input_template.xlsx', sheet_name='load_patterns')
        except:
            # 기본 패턴 생성
            load_patterns_df = pd.DataFrame({
                'hour': range(1, 8761),
                'pattern': np.random.uniform(0.5, 1.5, 8760)
            })
        
        # integrated_input_data.xlsx 생성
        output_file = "integrated_input_data.xlsx"
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            buses_df.to_excel(writer, sheet_name='buses', index=False)
            generators_df.to_excel(writer, sheet_name='generators', index=False)
            loads_df.to_excel(writer, sheet_name='loads', index=False)
            stores_df.to_excel(writer, sheet_name='stores', index=False)
            links_df.to_excel(writer, sheet_name='links', index=False)
            lines_df.to_excel(writer, sheet_name='lines', index=False)  # 지역간 연결 추가
            constraints_df.to_excel(writer, sheet_name='constraints', index=False)  # 제약조건 추가
            timeseries_df.to_excel(writer, sheet_name='timeseries', index=False)
            renewable_patterns_df.to_excel(writer, sheet_name='renewable_patterns', index=False)
            load_patterns_df.to_excel(writer, sheet_name='load_patterns', index=False)
        
        print(f"\n=== 통합 결과 ===")
        print(f"총 buses: {len(buses_df)}개")
        print(f"총 generators: {len(generators_df)}개")
        print(f"총 loads: {len(loads_df)}개")
        print(f"총 stores: {len(stores_df)}개")
        print(f"총 links: {len(links_df)}개")
        print(f"총 lines: {len(lines_df)}개")  # 지역간 연결
        print(f"총 constraints: {len(constraints_df)}개")  # 제약조건
        
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
    print("=== 수정된 integrated_input_data.xlsx 생성 (lines + constraints 포함) ===")
    create_integrated_data() 