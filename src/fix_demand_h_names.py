import pandas as pd
import os

# 입력 및 출력 파일 경로
input_file = 'integrated_input_data_fixed_loads.xlsx'  # 이전 스크립트의 출력을 입력으로 사용
output_file = 'integrated_input_data_fixed_all_loads.xlsx'
backup_file = 'integrated_input_data_fixed_loads_backup.xlsx'

def fix_demand_h_names():
    print("Demand_H 형식의 로드 이름 수정 중...")
    
    # 백업 생성
    if os.path.exists(input_file):
        try:
            # 원본 파일 백업
            import shutil
            shutil.copy2(input_file, backup_file)
            print(f"입력 파일을 {backup_file}으로 백업했습니다.")
            
            # 원본 파일에서 데이터 로드
            with pd.ExcelFile(input_file) as xls:
                # 모든 시트 로드
                data = {}
                for sheet_name in xls.sheet_names:
                    data[sheet_name] = pd.read_excel(xls, sheet_name)
            
            # _Demand_H 이름을 _Demand_H2로 변경
            if 'loads' in data:
                # _Demand_H로 끝나는 로드 찾기
                mask = data['loads']['name'].str.endswith('_Demand_H')
                h_demand_loads = data['loads'][mask].copy()
                
                if not h_demand_loads.empty:
                    print(f"수정할 로드 수: {len(h_demand_loads)}")
                    
                    # 각 로드에 대해 이름 변경
                    for idx, row in h_demand_loads.iterrows():
                        old_name = row['name']
                        new_name = old_name + '2'  # _Demand_H -> _Demand_H2
                        print(f"이름 변경: {old_name} -> {new_name}")
                        
                        # 데이터프레임에서 이름 업데이트
                        data['loads'].loc[data['loads']['name'] == old_name, 'name'] = new_name
                    
                    # 수정 후 올바른 형식 확인
                    correct_format = data['loads']['name'].str.contains('_Demand_EL$|_Demand_H2$', case=False)
                    incorrect_names = data['loads'][~correct_format]['name'].tolist()
                    
                    if incorrect_names:
                        print("\n여전히 형식이 맞지 않는 로드 이름들:")
                        for name in incorrect_names:
                            print(f"- {name}")
                    else:
                        print("\n모든 로드 이름이 올바른 형식을 따릅니다.")
                else:
                    print("수정할 _Demand_H 형식의 로드가 없습니다.")
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
    fix_demand_h_names() 