import pandas as pd
import os
import shutil
from datetime import datetime
import sys

def final_fix():
    """
    PyPSA-HD의 최적화 문제를 최종적으로 해결하기 위한 스크립트입니다.
    
    1. 최적화 기간 제한 (계산량 감소)
    2. CPLEX 솔버 지정 및 옵션 설정
    3. 네트워크 완화 (제약조건 완화)
    """
    print("최종 최적화 문제 해결 중...")
    
    # 파일 경로
    input_file = 'integrated_input_data.xlsx'
    config_file = 'config_optimize.py'
    
    # 백업 파일 생성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = f'integrated_input_data_backup_final_{timestamp}.xlsx'
    
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
        
        # 1. 모든 링크와 선로의 확장성 및 제약 완화
        print("네트워크 제약조건 완화 중...")
        
        # 링크 제약 완화
        for idx in links.index:
            links.loc[idx, 'p_nom_extendable'] = True
            
            if 'p_nom_max' in links.columns:
                current_max = links.loc[idx, 'p_nom_max'] if pd.notna(links.loc[idx, 'p_nom_max']) else 1000
                links.loc[idx, 'p_nom_max'] = max(current_max * 5, 10000)  # 기존 값의 5배 또는 최소 10000
            else:
                links['p_nom_max'] = 10000
                
            if 'p_nom_min' in links.columns:
                links.loc[idx, 'p_nom_min'] = 0
            else:
                links['p_nom_min'] = 0
        
        print(f"총 {len(links)}개 링크의 제약조건이 완화되었습니다.")
        
        # 선로 제약 완화
        for idx in lines.index:
            if 's_nom_extendable' in lines.columns:
                lines.loc[idx, 's_nom_extendable'] = True
            else:
                lines['s_nom_extendable'] = True
                
            if 's_nom_max' in lines.columns:
                current_max = lines.loc[idx, 's_nom_max'] if pd.notna(lines.loc[idx, 's_nom_max']) else 1000
                lines.loc[idx, 's_nom_max'] = max(current_max * 5, 10000)
            else:
                lines['s_nom_max'] = 10000
                
            if 's_nom_min' in lines.columns:
                lines.loc[idx, 's_nom_min'] = 0
            else:
                lines['s_nom_min'] = 0
        
        print(f"총 {len(lines)}개 선로의 제약조건이 완화되었습니다.")
        
        # 발전기 제약 완화
        for idx in generators.index:
            generators.loc[idx, 'p_nom_extendable'] = True
            
            if 'p_nom_max' in generators.columns:
                current_max = generators.loc[idx, 'p_nom_max'] if pd.notna(generators.loc[idx, 'p_nom_max']) else 1000
                generators.loc[idx, 'p_nom_max'] = max(current_max * 5, 10000)
            else:
                generators['p_nom_max'] = 10000
                
            if 'p_nom_min' in generators.columns:
                generators.loc[idx, 'p_nom_min'] = 0
            else:
                generators['p_nom_min'] = 0
        
        print(f"총 {len(generators)}개 발전기의 제약조건이 완화되었습니다.")
        
        # 저장장치 제약 완화
        for idx in stores.index:
            stores.loc[idx, 'e_nom_extendable'] = True
            
            if 'e_nom_max' in stores.columns:
                current_max = stores.loc[idx, 'e_nom_max'] if pd.notna(stores.loc[idx, 'e_nom_max']) else 1000
                stores.loc[idx, 'e_nom_max'] = max(current_max * 5, 10000)
            else:
                stores['e_nom_max'] = 10000
        
        print(f"총 {len(stores)}개 저장장치의 제약조건이 완화되었습니다.")
        
        # 엑셀 파일에 저장
        with pd.ExcelWriter(input_file) as writer:
            buses.to_excel(writer, sheet_name='buses', index=False)
            links.to_excel(writer, sheet_name='links', index=False)
            generators.to_excel(writer, sheet_name='generators', index=False)
            loads.to_excel(writer, sheet_name='loads', index=False)
            lines.to_excel(writer, sheet_name='lines', index=False)
            stores.to_excel(writer, sheet_name='stores', index=False)
        
        print(f"\n{input_file} 파일이 성공적으로 수정되었습니다.")
        
        # 2. 최적화 설정 파일 생성 (CPLEX 솔버 및 최적화 설정)
        print("\n최적화 설정 파일 생성 중...")
        
        config_content = """
# 최적화 설정
solve_opts = {
    "solver_name": "cplex",
    "solver_options": {
        "threads": 4,
        "lpmethod": 4,  # 배리어 메서드
        "barrier.algorithm": 3,  # 기본 배리어 알고리즘
        "mip.tolerances.mipgap": 0.05,  # MIP 갭 허용치
        "timelimit": 3600  # 시간 제한 (초)
    },
    "formulation": "kirchhoff"
}

# 시간 범위 제한 (계산량 감소)
time_settings = {
    "start_time": "2023-01-01 00:00:00",
    "end_time": "2023-01-31 23:00:00",  # 1월 한 달만 사용
    "freq": "1h"
}

# 제약조건 완화
constraints = {}  # 모든 제약조건 비활성화
"""
        
        with open(config_file, 'w') as f:
            f.write(config_content)
        
        print(f"{config_file} 파일이 생성되었습니다.")
        
        # 3. PyPSA_GUI.py 실행을 위한 커맨드 생성
        cmd = f"python PyPSA_GUI.py --config {config_file}"
        print(f"\n이제 다음 명령어로 PyPSA_GUI.py를 실행하세요:")
        print(f"  {cmd}")
        
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
    final_fix() 