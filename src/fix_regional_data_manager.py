import os
import shutil
from datetime import datetime

def fix_regional_data_manager():
    """
    regional_data_manager.py 파일의 문제를 해결하는 함수
    1. 중복된 REGION_TEMPLATES 선언 제거
    2. 필요한 메서드 추가 (initialize_region, merge_data)
    """
    print("regional_data_manager.py 파일 수정 중...")
    
    # 파일 경로
    target_file = 'regional_data_manager.py'
    backup_file = f'regional_data_manager_backup_full_{datetime.now().strftime("%Y%m%d_%H%M%S")}.py'
    
    if not os.path.exists(target_file):
        print(f"오류: {target_file} 파일을 찾을 수 없습니다.")
        return False
    
    # 백업 생성
    shutil.copy2(target_file, backup_file)
    print(f"원본 파일을 {backup_file}로 백업했습니다.")
    
    try:
        # 파일 내용 읽기
        with open(target_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        # 수정할 내용 준비
        new_lines = []
        in_class_def = False
        already_added_methods = False
        
        for i, line in enumerate(lines):
            # 클래스 정의 후에 있는 중복 REGION_TEMPLATES = {} 라인 제거
            if in_class_def and line.strip() == "REGION_TEMPLATES = {}":
                continue
                
            # 클래스 정의 시작 확인
            if "class RegionalDataManager" in line:
                in_class_def = True
                new_lines.append(line)
                continue
                
            # 클래스 마지막에 메서드 추가
            if in_class_def and not already_added_methods and i < len(lines) - 1:
                next_line = lines[i + 1].strip()
                # 클래스 끝 또는 파일 끝 감지
                if (not next_line or next_line.startswith("def ") or 
                    next_line.startswith("class ") or next_line.startswith("#")):
                    # 필요한 메서드 추가
                    new_lines.append(line)
                    new_lines.append("""    def initialize_region(self, region_code):
        \"\"\"지역 초기화
        
        Args:
            region_code (str): 지역 코드
        \"\"\"
        if region_code not in self.regional_data:
            self.regional_data[region_code] = {
                'buses': [],
                'generators': [],
                'loads': [],
                'lines': [],
                'stores': [],
                'links': []
            }
            
            # 템플릿에서 기본 데이터 로드 (있는 경우)
            if region_code in REGION_TEMPLATES:
                for component, items in REGION_TEMPLATES[region_code].items():
                    self.regional_data[region_code][component].extend(items)
    
    def add_component(self, region_code, component_type, data):
        \"\"\"구성요소 추가
        
        Args:
            region_code (str): 지역 코드
            component_type (str): 구성요소 유형 (buses, generators, loads, lines, stores, links)
            data (dict): 구성요소 데이터
        \"\"\"
        # 지역이 초기화되어 있는지 확인
        if region_code not in self.regional_data:
            self.initialize_region(region_code)
            
        # 유효한 구성요소 유형인지 확인
        if component_type not in self.regional_data[region_code]:
            print(f"경고: '{component_type}'은(는) 유효한 구성요소 유형이 아닙니다.")
            return
            
        # 데이터 추가
        self.regional_data[region_code][component_type].append(data)
    
    def add_connection(self, region1, region2, connection_data=None):
        \"\"\"지역간 연결 추가
        
        Args:
            region1 (str): 시작 지역 코드
            region2 (str): 도착 지역 코드
            connection_data (dict, optional): 연결 속성 데이터
        \"\"\"
        # 기본 연결 데이터
        conn_data = {
            'name': f"{region1}_{region2}",
            'bus0': f"{region1}_EL",
            'bus1': f"{region2}_EL",
            'carrier': 'AC',
            's_nom': 1000,
            'v_nom': 345,
            'length': 100,
            'x': 0.02,
            'r': 0.005
        }
        
        # 사용자 제공 데이터로 업데이트
        if connection_data:
            conn_data.update(connection_data)
            
        self.connections.append(conn_data)
    
    def merge_data(self):
        \"\"\"모든 지역 데이터 병합
        
        Returns:
            dict: 구성요소별로 병합된 데이터프레임
        \"\"\"
        import pandas as pd
        
        # 반환할 통합 데이터
        merged_data = {
            'buses': [],
            'generators': [],
            'loads': [],
            'lines': [],
            'stores': [],
            'links': []
        }
        
        # 각 지역의 데이터 병합
        for region_code, region_data in self.regional_data.items():
            for component, items in region_data.items():
                if component in merged_data:
                    merged_data[component].extend(items)
        
        # 지역간 연결 추가
        merged_data['lines'].extend(self.connections)
        
        # 리스트를 데이터프레임으로 변환
        result = {}
        for component, items in merged_data.items():
            if items:  # 비어있지 않은 경우만
                df = pd.DataFrame(items)
                # 필요한 템플릿 컬럼 추가
                if component in self.templates:
                    required_columns = self.templates[component]['columns']
                    for col in required_columns:
                        if col not in df.columns:
                            df[col] = None
                    # 주요 컬럼 순서대로 정렬
                    df = df[required_columns + [c for c in df.columns if c not in required_columns]]
                result[component] = df
            else:
                result[component] = pd.DataFrame()
                
        return result
        
    def export_merged_data(self, output_path):
        \"\"\"병합된 데이터를 엑셀 파일로 내보내기
        
        Args:
            output_path (str): 출력 파일 경로
        \"\"\"
        merged_data = self.merge_data()
        
        with pd.ExcelWriter(output_path) as writer:
            for component, df in merged_data.items():
                if not df.empty:
                    df.to_excel(writer, sheet_name=component, index=False)
                    
        print(f"통합 데이터가 '{output_path}'에 저장되었습니다.")
""")
                    already_added_methods = True
                    continue
            
            # 모든 다른 라인 유지
            new_lines.append(line)
        
        # 수정된 내용 저장
        with open(target_file, 'w', encoding='utf-8') as f:
            f.writelines(new_lines)
        
        print(f"\n{target_file} 파일이 성공적으로 수정되었습니다.")
        print("이제 PyPSA_GUI.py 또는 PyPSA_HD_Regional.py를 다시 실행해보세요.")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        
        # 백업에서 복원
        print(f"백업에서 복원 중...")
        shutil.copy2(backup_file, target_file)
        print(f"원본 파일이 백업에서 복원되었습니다.")
        
        return False

if __name__ == "__main__":
    fix_regional_data_manager() 