import pandas as pd
import numpy as np
from datetime import datetime

def restore_chp_links():
    """백업에서 CHP 링크를 복원하는 함수"""
    
    print("=== CHP 링크 복원 ===")
    
    backup_file = 'integrated_input_data_backup_20250915_114842.xlsx'
    
    try:
        # 백업에서 링크 데이터 로드
        backup_links = pd.read_excel(backup_file, sheet_name='links')
        print(f"백업에서 {len(backup_links)}개 링크 로드됨")
        
        # 현재 파일에서 기존 데이터 로드
        current_data = {}
        with pd.ExcelFile('integrated_input_data.xlsx') as xls:
            for sheet in xls.sheet_names:
                if sheet != 'links':  # links는 제외하고 로드
                    current_data[sheet] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet)
                    print(f"{sheet} 시트 로드: {len(current_data[sheet])}개 행")
        
        # CHP 링크 확인
        chp_links = backup_links[backup_links['name'].str.contains('CHP', na=False)]
        print(f"\n복원할 CHP 링크: {len(chp_links)}개")
        
        total_chp_capacity = chp_links['p_nom'].sum()
        print(f"총 CHP 용량: {total_chp_capacity:.0f} MW")
        
        # CHP 링크 상세 출력
        print("\nCHP 링크 상세:")
        for _, link in chp_links.iterrows():
            name = link['name']
            bus0 = link['bus0']
            bus1 = link['bus1'] 
            bus2 = link.get('bus2', 'N/A')
            p_nom = link['p_nom']
            eff1 = link.get('efficiency1', 'N/A')
            eff2 = link.get('efficiency2', 'N/A')
            print(f"- {name}: {bus0}(LNG) -> {bus1}(EL) + {bus2}(H) | {p_nom}MW (전력eff:{eff1}, 열eff:{eff2})")
        
        # 수정된 파일 저장
        print(f"\n링크 시트 복원 중...")
        
        # 현재 파일 백업
        current_backup = f"integrated_input_data_before_chp_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        import shutil
        shutil.copy2('integrated_input_data.xlsx', current_backup)
        print(f"현재 파일 백업: {current_backup}")
        
        # 새 파일에 모든 시트 저장 (links 포함)
        with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
            # 백업에서 복원한 links 시트
            backup_links.to_excel(writer, sheet_name='links', index=False)
            
            # 기존 시트들
            for sheet_name, df in current_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n=== 복원 완료 ===")
        print(f"- CHP 링크 {len(chp_links)}개 복원됨")
        print(f"- 총 CHP 용량: {total_chp_capacity:.0f} MW")
        print("- 이제 백업 LNG 발전기 대신 CHP가 우선 사용될 것입니다")
        
        # 복원 후 확인
        restored_links = pd.read_excel('integrated_input_data.xlsx', sheet_name='links')
        print(f"\n복원 확인: {len(restored_links)}개 링크가 파일에 저장됨")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    restore_chp_links() 