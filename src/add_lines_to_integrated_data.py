import pandas as pd
import os

def add_lines_to_integrated_data():
    """
    지역간 연결 시트의 데이터를 integrated_input_data.xlsx 파일의 lines 시트로 추가합니다.
    """
    try:
        # 파일 경로 설정
        template_file = 'regional_input_template.xlsx'
        integrated_file = 'integrated_input_data.xlsx'
        
        # 파일 존재 확인
        if not os.path.exists(template_file):
            print(f"오류: '{template_file}' 파일을 찾을 수 없습니다.")
            return False
        
        if not os.path.exists(integrated_file):
            print(f"오류: '{integrated_file}' 파일을 찾을 수 없습니다.")
            return False
        
        # 지역간 연결 시트 읽기 - 헤더 행이 5행(0-based로는 4)에 있음
        connections_df = pd.read_excel(template_file, sheet_name='지역간 연결', header=4)
        
        if connections_df.empty:
            print("경고: '지역간 연결' 시트에 데이터가 없습니다.")
            return False
        
        # 연결 데이터를 lines 형식으로 변환
        lines_data = []
        for _, row in connections_df.iterrows():
            # 빈 행이거나 필수 데이터가 없는 경우 건너뛰기
            if pd.isna(row.iloc[0]) or row.iloc[0] is None or str(row.iloc[0]).strip() == '':
                continue
                
            # 컬럼 이름 확인
            name = row.iloc[0]       # 이름
            region1 = row.iloc[1]    # 시작 지역
            bus1 = row.iloc[2]       # 시작 버스
            region2 = row.iloc[3]    # 도착 지역
            bus2 = row.iloc[4]       # 도착 버스
            capacity = row.iloc[5]   # 정격용량(MVA)
            voltage = row.iloc[6]    # 전압(kV)
            distance = row.iloc[7]   # 거리(km)
            x = row.iloc[8]          # 리액턴스(p.u.)
            r = row.iloc[9]          # 저항(p.u.)
            
            # 필수 필드 확인
            if pd.isna(region1) or pd.isna(region2) or pd.isna(bus1) or pd.isna(bus2):
                continue
            
            # 버스 이름 확인 및 처리
            bus0 = bus1
            bus1 = bus2
            
            # 지역 접두사가 없는 경우 추가
            if bus0 and not str(bus0).startswith(f"{region1}_"):
                bus0 = f"{region1}_{bus0}"
            
            if bus1 and not str(bus1).startswith(f"{region2}_"):
                bus1 = f"{region2}_{bus1}"
            
            # 필수 값 설정
            line_data = {
                'name': name,
                'bus0': bus0,
                'bus1': bus1,
                'carrier': 'AC',
                's_nom': float(capacity) if pd.notna(capacity) else 1000.0,
                'v_nom': float(voltage) if pd.notna(voltage) else 345.0,
                'length': float(distance) if pd.notna(distance) else None,
                'x': float(x) if pd.notna(x) else None,
                'r': float(r) if pd.notna(r) else None
            }
            
            # 거리 기반 계산 (값이 없는 경우)
            if line_data['length'] is None or line_data['x'] is None or line_data['r'] is None:
                # 계산된 거리가 없으면 기본값 사용
                if line_data['length'] is None:
                    line_data['length'] = 100.0
                
                if line_data['x'] is None:
                    line_data['x'] = line_data['length'] * 0.0004
                
                if line_data['r'] is None:
                    line_data['r'] = line_data['length'] * 0.0001
            
            lines_data.append(line_data)
        
        if not lines_data:
            print("경고: 변환할 연결 데이터가 없습니다.")
            return False
        
        # 통합 데이터 파일 읽기
        xls = pd.ExcelFile(integrated_file)
        data_dict = {}
        
        # 기존 시트 읽기
        for sheet_name in xls.sheet_names:
            data_dict[sheet_name] = pd.read_excel(integrated_file, sheet_name=sheet_name)
        
        # lines 데이터 추가 또는 업데이트
        lines_df = pd.DataFrame(lines_data)
        data_dict['lines'] = lines_df
        
        # 파일로 저장
        with pd.ExcelWriter(integrated_file, engine='openpyxl') as writer:
            for sheet_name, df in data_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"'{template_file}'의 '지역간 연결' 시트에서 {len(lines_data)}개의 선로 데이터를 '{integrated_file}'의 'lines' 시트로 추가했습니다.")
        return True
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    add_lines_to_integrated_data() 