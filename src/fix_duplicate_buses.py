import pandas as pd
import os

# 입력 및 출력 파일 경로
input_file = 'integrated_input_data.xlsx'
output_file = 'integrated_input_data_fixed_all.xlsx'
backup_file = 'integrated_input_data_backup_all.xlsx'

def fix_duplicate_buses():
    print("중복된 버스 문제 해결 중...")
    
    # 백업 생성
    if os.path.exists(input_file):
        try:
            # 이미 수정된 파일이 있다면 그것을 사용
            if os.path.exists('integrated_input_data_fixed_buses.xlsx'):
                print("이전에 수정된 파일(integrated_input_data_fixed_buses.xlsx)을 사용합니다.")
                with pd.ExcelFile('integrated_input_data_fixed_buses.xlsx') as xls:
                    # 모든 시트 로드
                    data = {}
                    for sheet_name in xls.sheet_names:
                        data[sheet_name] = pd.read_excel(xls, sheet_name)
            else:
                # 원본 파일 백업 및 사용
                import shutil
                shutil.copy2(input_file, backup_file)
                print(f"원본 파일을 {backup_file}으로 백업했습니다.")
                
                # 원본 파일에서 데이터 로드
                with pd.ExcelFile(input_file) as xls:
                    # 모든 시트 로드
                    data = {}
                    for sheet_name in xls.sheet_names:
                        data[sheet_name] = pd.read_excel(xls, sheet_name)
            
            # 1. 전력 버스 중복 처리 - JBD_Main_EL과 SEL_Main_EL을 JBD_EL와 SEL_EL로 변경
            print("\n1. 전력 버스 중복 처리 중...")
            
            # 버스 변경 매핑
            el_bus_mapping = {'JBD_Main_EL': 'JBD_EL', 'SEL_Main_EL': 'SEL_EL'}
            
            # buses 시트에서 중복 전력 버스 제거
            problematic_el_buses = list(el_bus_mapping.keys())
            data['buses'] = data['buses'][~data['buses']['name'].isin(problematic_el_buses)]
            print(f"bus 시트에서 {problematic_el_buses} 버스를 제거했습니다.")
            
            # loads 시트에서 중복 전력 버스로 향하는 부하 처리
            for old_bus, new_bus in el_bus_mapping.items():
                mask = data['loads']['bus'] == old_bus
                if mask.any():
                    # 기존의 같은 이름의 전력 버스로 가는 부하가 있는지 확인
                    duplicate_loads = data['loads'][(data['loads']['bus'] == new_bus) & 
                                                  (data['loads']['name'].str.contains('Demand_EL', na=False))]
                    
                    # 기존 부하가 있으면 해당 부하의 p_set 값을 가져와서 더함
                    for _, load_row in data['loads'][mask].iterrows():
                        duplicate_mask = (data['loads']['bus'] == new_bus) & (data['loads']['name'] == f"{new_bus.split('_')[0]}_Demand_EL")
                        
                        if duplicate_mask.any():
                            # 기존 부하의 p_set 값에 추가 (중복 부하는 하나씩만 있다고 가정)
                            original_p_set = data['loads'].loc[duplicate_mask, 'p_set'].values[0]
                            data['loads'].loc[duplicate_mask, 'p_set'] += load_row['p_set']
                            print(f"{new_bus} 버스의 기존 부하 p_set 값을 {original_p_set}에서 {data['loads'].loc[duplicate_mask, 'p_set'].values[0]}로 증가시켰습니다.")
                        else:
                            # 기존 부하가 없으면 새 부하를 추가하지 않고 이름만 변경
                            row_idx = data['loads'][mask].index
                            data['loads'].loc[row_idx, 'bus'] = new_bus
                            data['loads'].loc[row_idx, 'name'] = f"{new_bus.split('_')[0]}_Demand_EL"
                            data['loads'].loc[row_idx, 'carrier'] = 'electricity'
                            print(f"{old_bus} 버스로 가는 부하를 {new_bus} 버스로 변경했습니다.")
                    
                    # 중복 부하 제거
                    data['loads'] = data['loads'][~mask]
            
            # 2. 수소 버스 중복 처리 - JBD_Hydrogen과 SEL_Hydrogen을 JBD_H2와 SEL_H2로 변경
            print("\n2. 수소 버스 중복 처리 중...")
            
            # 버스 변경 매핑
            h2_bus_mapping = {'JBD_Hydrogen': 'JBD_H2', 'SEL_Hydrogen': 'SEL_H2'}
            
            # buses 시트에서 중복 수소 버스 제거
            problematic_h2_buses = list(h2_bus_mapping.keys())
            data['buses'] = data['buses'][~data['buses']['name'].isin(problematic_h2_buses)]
            print(f"bus 시트에서 {problematic_h2_buses} 버스를 제거했습니다.")
            
            # loads 시트에서 중복 수소 버스로 향하는 부하 처리
            for old_bus, new_bus in h2_bus_mapping.items():
                mask = data['loads']['bus'] == old_bus
                if mask.any():
                    # 기존의 같은 이름의 수소 버스로 가는 부하가 있는지 확인
                    duplicate_loads = data['loads'][(data['loads']['bus'] == new_bus) & 
                                                  (data['loads']['name'].str.contains('Demand_H2', na=False))]
                    
                    # 기존 부하가 있으면 해당 부하의 p_set 값을 가져와서 더함
                    for _, load_row in data['loads'][mask].iterrows():
                        duplicate_mask = (data['loads']['bus'] == new_bus) & (data['loads']['name'] == f"{new_bus.split('_')[0]}_Demand_H2")
                        
                        if duplicate_mask.any():
                            # 기존 부하의 p_set 값에 추가 (중복 부하는 하나씩만 있다고 가정)
                            original_p_set = data['loads'].loc[duplicate_mask, 'p_set'].values[0]
                            data['loads'].loc[duplicate_mask, 'p_set'] += load_row['p_set']
                            print(f"{new_bus} 버스의 기존 부하 p_set 값을 {original_p_set}에서 {data['loads'].loc[duplicate_mask, 'p_set'].values[0]}로 증가시켰습니다.")
                        else:
                            # 기존 부하가 없으면 새 부하를 추가하지 않고 이름만 변경
                            row_idx = data['loads'][mask].index
                            data['loads'].loc[row_idx, 'bus'] = new_bus
                            data['loads'].loc[row_idx, 'name'] = f"{new_bus.split('_')[0]}_Demand_H2"
                            data['loads'].loc[row_idx, 'carrier'] = 'hydrogen'
                            print(f"{old_bus} 버스로 가는 부하를 {new_bus} 버스로 변경했습니다.")
                    
                    # 중복 부하 제거
                    data['loads'] = data['loads'][~mask]
            
            # 3. 다른 시트에서도 버스 이름 변경
            print("\n3. 다른 시트에서 버스 이름 변경 중...")
            
            # 모든 버스 변경 매핑 통합
            all_bus_mapping = {**el_bus_mapping, **h2_bus_mapping}
            
            # generators, links, stores 시트에서 버스 이름 변경
            for component in ['generators', 'links', 'stores']:
                if component in data:
                    for old_bus, new_bus in all_bus_mapping.items():
                        # bus0 필드가 있는 경우 (links)
                        if 'bus0' in data[component].columns:
                            mask = data[component]['bus0'] == old_bus
                            if mask.any():
                                data[component].loc[mask, 'bus0'] = new_bus
                                print(f"{component} 시트에서 bus0 필드의 {old_bus}를 {new_bus}로 변경했습니다.")
                        
                        # bus1 필드가 있는 경우 (links)
                        if 'bus1' in data[component].columns:
                            mask = data[component]['bus1'] == old_bus
                            if mask.any():
                                data[component].loc[mask, 'bus1'] = new_bus
                                print(f"{component} 시트에서 bus1 필드의 {old_bus}를 {new_bus}로 변경했습니다.")
                        
                        # bus 필드가 있는 경우 (generators, stores)
                        if 'bus' in data[component].columns:
                            mask = data[component]['bus'] == old_bus
                            if mask.any():
                                data[component].loc[mask, 'bus'] = new_bus
                                print(f"{component} 시트에서 bus 필드의 {old_bus}를 {new_bus}로 변경했습니다.")
            
            # 결과 저장
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for sheet_name, df in data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"\n수정된 데이터가 {output_file}에 저장되었습니다.")
        except Exception as e:
            print(f"오류 발생: {e}")
    else:
        print(f"오류: {input_file} 파일이 존재하지 않습니다.")

if __name__ == "__main__":
    fix_duplicate_buses() 