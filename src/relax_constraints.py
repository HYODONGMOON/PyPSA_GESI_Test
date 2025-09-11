import pandas as pd
import os
import shutil
from datetime import datetime

def relax_constraints(input_file='simplified_input_data.xlsx'):
    """
    PyPSA-HD 모델의 제약조건을 완화하는 함수
    """
    print(f"'{input_file}' 파일의 제약조건 완화 중...")
    
    # 파일 존재 확인
    if not os.path.exists(input_file):
        print(f"오류: '{input_file}' 파일을 찾을 수 없습니다.")
        return False
    
    # 백업 파일 생성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = f'{os.path.splitext(input_file)[0]}_backup_{timestamp}.xlsx'
    shutil.copy2(input_file, backup_file)
    print(f"원본 파일을 '{backup_file}'로 백업했습니다.")
    
    try:
        # 파일 로드
        with pd.ExcelFile(input_file) as xls:
            buses = pd.read_excel(xls, sheet_name='buses')
            links = pd.read_excel(xls, sheet_name='links')
            generators = pd.read_excel(xls, sheet_name='generators')
            loads = pd.read_excel(xls, sheet_name='loads')
            lines = pd.read_excel(xls, sheet_name='lines')
            stores = pd.read_excel(xls, sheet_name='stores')
        
        print("데이터 로드 완료. 제약조건 완화 중...")
        
        # 링크 제약 완화
        for idx in links.index:
            links.loc[idx, 'p_nom_extendable'] = True
            if 'p_nom_max' in links.columns:
                links.loc[idx, 'p_nom_max'] = 10000
            else:
                links['p_nom_max'] = 10000
            if 'p_nom_min' in links.columns:
                links.loc[idx, 'p_nom_min'] = 0
            else:
                links['p_nom_min'] = 0
        
        print(f"링크 제약조건 완화 완료: {len(links)}개")
        
        # 선로 제약 완화
        for idx in lines.index:
            if 's_nom_extendable' in lines.columns:
                lines.loc[idx, 's_nom_extendable'] = True
            else:
                lines['s_nom_extendable'] = True
            if 's_nom_max' in lines.columns:
                lines.loc[idx, 's_nom_max'] = 10000
            else:
                lines['s_nom_max'] = 10000
            if 's_nom_min' in lines.columns:
                lines.loc[idx, 's_nom_min'] = 0
            else:
                lines['s_nom_min'] = 0
        
        print(f"선로 제약조건 완화 완료: {len(lines)}개")
        
        # 발전기 제약 완화
        for idx in generators.index:
            generators.loc[idx, 'p_nom_extendable'] = True
            if 'p_nom_max' in generators.columns:
                generators.loc[idx, 'p_nom_max'] = 10000
            else:
                generators['p_nom_max'] = 10000
            if 'p_nom_min' in generators.columns:
                generators.loc[idx, 'p_nom_min'] = 0
            else:
                generators['p_nom_min'] = 0
        
        print(f"발전기 제약조건 완화 완료: {len(generators)}개")
        
        # 저장장치 제약 완화
        for idx in stores.index:
            stores.loc[idx, 'e_nom_extendable'] = True
            if 'e_nom_max' in stores.columns:
                stores.loc[idx, 'e_nom_max'] = 10000
            else:
                stores['e_nom_max'] = 10000
            if 'standing_loss' in stores.columns and pd.isna(stores.loc[idx, 'standing_loss']):
                stores.loc[idx, 'standing_loss'] = 0.0
        
        print(f"저장장치 제약조건 완화 완료: {len(stores)}개")
        
        # carrier 속성 확인 및 수정
        if 'carrier' in buses.columns:
            for idx in buses.index:
                if pd.isna(buses.loc[idx, 'carrier']):
                    buses.loc[idx, 'carrier'] = 'AC'
        
        if 'carrier' in generators.columns:
            for idx in generators.index:
                if pd.isna(generators.loc[idx, 'carrier']):
                    generators.loc[idx, 'carrier'] = 'AC'
        
        if 'carrier' in lines.columns:
            for idx in lines.index:
                if pd.isna(lines.loc[idx, 'carrier']):
                    lines.loc[idx, 'carrier'] = 'AC'
        
        print("carrier 속성 확인 및 수정 완료")
        
        # 수정된 파일 저장
        with pd.ExcelWriter(input_file) as writer:
            buses.to_excel(writer, sheet_name='buses', index=False)
            links.to_excel(writer, sheet_name='links', index=False)
            generators.to_excel(writer, sheet_name='generators', index=False)
            loads.to_excel(writer, sheet_name='loads', index=False)
            lines.to_excel(writer, sheet_name='lines', index=False)
            stores.to_excel(writer, sheet_name='stores', index=False)
        
        print(f"\n제약조건이 완화된 파일 '{input_file}'이 저장되었습니다.")
        
        # 최적화 명령어 안내
        print("\n다음 명령어로 최적화를 실행하세요:")
        print(f"python run_cplex_optim.py --input {input_file}")
        
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
    import argparse
    
    parser = argparse.ArgumentParser(description='PyPSA-HD 모델의 제약조건을 완화합니다.')
    parser.add_argument('--input', default='simplified_input_data.xlsx', 
                        help='입력 파일 경로 (기본값: simplified_input_data.xlsx)')
    args = parser.parse_args()
    
    relax_constraints(args.input) 