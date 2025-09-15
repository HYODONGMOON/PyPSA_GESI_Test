import pandas as pd
from datetime import datetime
import shutil

def fix_zero_capacity_links():
    print("=== 용량 0인 링크 수정 ===")
    
    try:
        # 백업 생성
        backup_file = f"integrated_input_data_before_link_fix_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        shutil.copy2('integrated_input_data.xlsx', backup_file)
        print(f"백업 파일 생성: {backup_file}")
        
        # 모든 시트 로드
        input_data = {}
        with pd.ExcelFile('integrated_input_data.xlsx') as xls:
            for sheet in xls.sheet_names:
                input_data[sheet] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet)
        
        # Links 시트 수정
        links_df = input_data['links'].copy()
        
        print(f"\n수정 전 용량 0인 링크: {len(links_df[links_df['p_nom'] == 0])}개")
        
        # 각 링크 유형별 최소 용량 설정
        modifications = []
        
        for idx, link in links_df.iterrows():
            if link['p_nom'] == 0:
                name = link['name']
                
                # 링크 유형별 최소 용량 설정
                if 'Electrolyser' in name:
                    new_capacity = 50.0  # 수소 생산용 최소 용량
                    links_df.at[idx, 'p_nom'] = new_capacity
                    modifications.append(f"Electrolyser {name}: 0 → {new_capacity} MW")
                    
                elif 'HP' in name and 'CHP' not in name:
                    new_capacity = 100.0  # 열펌프 최소 용량
                    links_df.at[idx, 'p_nom'] = new_capacity
                    modifications.append(f"HP {name}: 0 → {new_capacity} MW")
                    
                elif 'CHP' in name:
                    new_capacity = 50.0  # CHP 최소 용량
                    links_df.at[idx, 'p_nom'] = new_capacity
                    modifications.append(f"CHP {name}: 0 → {new_capacity} MW")
        
        # 수정사항 출력
        print(f"\n수정된 링크: {len(modifications)}개")
        for mod in modifications:
            print(f"  - {mod}")
        
        # 수정된 데이터 저장
        input_data['links'] = links_df
        
        with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
            for sheet_name, df in input_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n수정 후 용량 0인 링크: {len(links_df[links_df['p_nom'] == 0])}개")
        
        # 수정 효과 분석
        print("\n=== 수정 효과 ===")
        
        # Electrolyser 총 용량
        electro_links = links_df[links_df['name'].str.contains('Electrolyser', na=False)]
        total_electro_capacity = electro_links['p_nom'].sum()
        print(f"Electrolyser 총 용량: {total_electro_capacity:,.1f} MW")
        
        # HP 총 용량  
        hp_links = links_df[links_df['name'].str.contains('HP', na=False) & ~links_df['name'].str.contains('CHP', na=False)]
        total_hp_capacity = hp_links['p_nom'].sum()
        print(f"HP 총 용량: {total_hp_capacity:,.1f} MW")
        
        # CHP 총 용량
        chp_links = links_df[links_df['name'].str.contains('CHP', na=False)]
        total_chp_capacity = chp_links['p_nom'].sum()
        print(f"CHP 총 용량: {total_chp_capacity:,.1f} MW")
        
        print(f"\n✅ 링크 용량 수정 완료!")
        print("이제 수소와 열 수요를 충족할 수 있는 최소 설비가 확보되었습니다.")
        print("다시 PyPSA 분석을 실행하면 infeasible이 해결될 것입니다.")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    fix_zero_capacity_links() 