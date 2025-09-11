import pandas as pd
import os

# 입력 및 출력 파일 경로
input_file = 'integrated_input_data.xlsx'
output_file = 'integrated_input_data_fixed_loads.xlsx'
backup_file = 'integrated_input_data_backup_loads.xlsx'

def fix_load_names():
    print("로드 이름 포맷 문제 해결 중...")
    
    # 백업 생성
    if os.path.exists(input_file):
        try:
            # 원본 파일 백업
            import shutil
            shutil.copy2(input_file, backup_file)
            print(f"원본 파일을 {backup_file}으로 백업했습니다.")
            
            # 원본 파일에서 데이터 로드
            with pd.ExcelFile(input_file) as xls:
                # 모든 시트 로드
                data = {}
                for sheet_name in xls.sheet_names:
                    data[sheet_name] = pd.read_excel(xls, sheet_name)
            
            # 지정된 로드만 제거 (JBD_H2_Demand, JBD_Demand1, SEL_Demand1, SEL_H2_Demand)
            if 'loads' in data:
                # 제거할 로드 이름 목록
                loads_to_remove = ['JBD_H2_Demand', 'JBD_Demand1', 'SEL_Demand1', 'SEL_H2_Demand']
                
                # 제거하기 전 정보 출력
                removed_loads = data['loads'][data['loads']['name'].isin(loads_to_remove)]
                print(f"제거될 로드 수: {len(removed_loads)}")
                for idx, row in removed_loads.iterrows():
                    print(f"제거: {row['name']} (버스: {row['bus']}, p_set: {row['p_set']})")
                
                # 로드 제거
                data['loads'] = data['loads'][~data['loads']['name'].isin(loads_to_remove)]
                
                # 남은 로드 이름 형식 확인
                print("\n남은 로드 형식 확인:")
                for name in data['loads']['name'].unique():
                    print(f"- {name}")
                
                # 올바른 형식인지 확인 (지역_Demand_EL, 지역_Demand_H, 지역_Demand_H2 형식)
                correct_format = data['loads']['name'].str.contains('_Demand_EL$|_Demand_H$|_Demand_H2$', case=False)
                incorrect_names = data['loads'][~correct_format]['name'].tolist()
                
                if incorrect_names:
                    print("\n여전히 형식이 맞지 않는 로드 이름들:")
                    for name in incorrect_names:
                        print(f"- {name}")
                else:
                    print("\n모든 로드 이름이 올바른 형식을 따릅니다 (_Demand_EL, _Demand_H, _Demand_H2).")
                
                # 각 형식별 로드 수 확인
                el_loads = data['loads'][data['loads']['name'].str.contains('_Demand_EL$', case=False)]
                h_loads = data['loads'][data['loads']['name'].str.contains('_Demand_H$', case=False)]
                h2_loads = data['loads'][data['loads']['name'].str.contains('_Demand_H2$', case=False)]
                
                print(f"\n전력 로드 (_Demand_EL) 개수: {len(el_loads)}")
                print(f"열 로드 (_Demand_H) 개수: {len(h_loads)}")
                print(f"수소 로드 (_Demand_H2) 개수: {len(h2_loads)}")
            else:
                print("loads 시트가 존재하지 않습니다.")
            
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
    fix_load_names() 