import pandas as pd
import os
import shutil

# 파일 경로
input_file = 'integrated_input_data.xlsx'
backup_file = 'integrated_input_data_backup.xlsx'

def fix_and_overwrite():
    print("문제 해결 및 파일 수정 중...")
    
    # 백업 생성
    if os.path.exists(input_file):
        try:
            # 원본 파일 백업
            shutil.copy2(input_file, backup_file)
            print(f"원본 파일을 {backup_file}으로 백업했습니다.")
            
            # 원본 파일에서 데이터 로드
            with pd.ExcelFile(input_file) as xls:
                # 모든 시트 로드
                data = {}
                for sheet_name in xls.sheet_names:
                    data[sheet_name] = pd.read_excel(xls, sheet_name)
            
            # 1. 지정된 로드 항목 제거
            if 'loads' in data:
                # 제거할 로드 이름 목록
                loads_to_remove = ['JBD_H2_Demand', 'JBD_Demand1', 'SEL_Demand1', 'SEL_H2_Demand']
                
                # 제거하기 전 정보 출력
                removed_loads = data['loads'][data['loads']['name'].isin(loads_to_remove)]
                print(f"\n[1] 로드 제거: 제거될 로드 수: {len(removed_loads)}")
                for idx, row in removed_loads.iterrows():
                    print(f"제거: {row['name']} (버스: {row['bus']}, p_set: {row['p_set']})")
                
                # 로드 제거
                data['loads'] = data['loads'][~data['loads']['name'].isin(loads_to_remove)]
                
                # 수정 후 로드 형식 확인
                el_loads = data['loads'][data['loads']['name'].str.contains('_Demand_EL$', case=False)]
                h_loads = data['loads'][data['loads']['name'].str.contains('_Demand_H$', case=False)]
                h2_loads = data['loads'][data['loads']['name'].str.contains('_Demand_H2$', case=False)]
                
                print(f"\n수정 후 타입별 로드 개수:")
                print(f"전력 로드 (_Demand_EL) 개수: {len(el_loads)}")
                print(f"열 로드 (_Demand_H) 개수: {len(h_loads)}")
                print(f"수소 로드 (_Demand_H2) 개수: {len(h2_loads)}")
            
            # 2. JJD_JND 라인의 리액턴스(x) 값 확인 및 수정
            if 'lines' in data:
                # JJD_JND 라인 선택
                mask = data['lines']['name'] == 'JJD_JND'
                if mask.any():
                    # 리액턴스 값 확인
                    x_value = data['lines'].loc[mask, 'x'].values[0]
                    r_value = data['lines'].loc[mask, 'r'].values[0]
                    
                    print(f"\n[2] 라인 리액턴스 확인: JJD_JND 라인의 현재 x 값: {x_value}, r 값: {r_value}")
                    
                    # 리액턴스 값이 0이면 수정
                    if x_value == 0:
                        new_x_value = r_value * 10  # r 값의 10배 정도로 설정
                        data['lines'].loc[mask, 'x'] = new_x_value
                        print(f"JJD_JND 라인의 x 값을 {x_value}에서 {new_x_value}로 변경했습니다.")
                    else:
                        print("JJD_JND 라인의 x 값은 이미 0이 아닙니다. 수정이 필요하지 않습니다.")
                else:
                    print("\n[2] 라인 리액턴스 확인: JJD_JND 라인을 찾을 수 없습니다.")
            
            # 3. 수소 버스에 연결된 로드 확인
            h2_buses = ['JBD_Hydrogen', 'SEL_Hydrogen']
            h2_bus_loads = data['loads'][data['loads']['bus'].isin(h2_buses)]
            
            print(f"\n[3] 수소 버스 로드 확인: {len(h2_bus_loads)}개 로드 발견")
            if not h2_bus_loads.empty:
                print("수소 버스에 연결된 로드:")
                for idx, row in h2_bus_loads.iterrows():
                    print(f"- {row['name']} (버스: {row['bus']}, p_set: {row['p_set']})")
            else:
                print("수소 버스에 연결된 로드가 없습니다. 문제가 해결되었습니다.")
            
            # 결과 저장 (원본 파일에 덮어쓰기)
            with pd.ExcelWriter(input_file, engine='openpyxl') as writer:
                for sheet_name, df in data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"\n모든 수정이 완료되었습니다. 원본 파일 '{input_file}'이 수정되었습니다.")
            print(f"백업 파일은 '{backup_file}'에 저장되어 있습니다.")
        except Exception as e:
            print(f"오류 발생: {e}")
    else:
        print(f"오류: {input_file} 파일이 존재하지 않습니다.")

if __name__ == "__main__":
    fix_and_overwrite() 