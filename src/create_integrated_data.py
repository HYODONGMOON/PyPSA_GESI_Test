import pandas as pd
import numpy as np
import os
from datetime import datetime
from create_integrated_data_v3 import create_integrated_data as v3_create_integrated_data
INTERFACE_XLSX = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'interface.xlsx'))

def get_selected_regions():
    """지역 선택 시트에서 선택된 지역 목록 가져오기"""
    try:
        df = pd.read_excel(INTERFACE_XLSX, sheet_name='지역 선택')
        
        # 헤더 찾기
        header_row = None
        for i, row in df.iterrows():
            if '코드' in str(row.iloc[0]) and '지역명' in str(row.iloc[1]):
                header_row = i
                break
        
        if header_row is None:
            print("헤더를 찾을 수 없습니다.")
            return []
        
        # 헤더 이후 데이터 읽기
        data_df = df.iloc[header_row+1:].copy()
        data_df.columns = ['코드', '지역명', '영문명', '선택', '인구', '면적']
        
        # 선택된 지역 필터링
        selected_regions = []
        for _, row in data_df.iterrows():
            if pd.notna(row['코드']) and str(row['선택']).strip().upper() == 'O':
                selected_regions.append(str(row['코드']).strip())
        
        print(f"선택된 지역: {selected_regions}")
        return selected_regions
        
    except Exception as e:
        print(f"지역 선택 읽기 오류: {str(e)}")
        return []

def read_regional_data(region_code):
    """특정 지역의 데이터 읽기"""
    try:
        sheet_name = f'지역_{region_code}'
        df = pd.read_excel(INTERFACE_XLSX, sheet_name=sheet_name)
        
        # 데이터 구조 파악
        data_sections = {}
        current_section = None
        
        for i, row in df.iterrows():
            first_col = str(row.iloc[0]).strip()
            
            # 섹션 헤더 찾기
            if '버스' in first_col and 'buses' not in data_sections:
                current_section = 'buses'
                data_sections[current_section] = {'start': i, 'data': []}
            elif '발전기' in first_col and 'generators' not in data_sections:
                current_section = 'generators'
                data_sections[current_section] = {'start': i, 'data': []}
            elif '부하' in first_col and 'loads' not in data_sections:
                current_section = 'loads'
                data_sections[current_section] = {'start': i, 'data': []}
            elif '저장장치' in first_col and 'stores' not in data_sections:
                current_section = 'stores'
                data_sections[current_section] = {'start': i, 'data': []}
            elif '링크' in first_col and 'links' not in data_sections:
                current_section = 'links'
                data_sections[current_section] = {'start': i, 'data': []}
            elif current_section and pd.notna(row.iloc[0]) and first_col not in ['', 'NaN']:
                # 데이터 행인지 확인
                if not any(keyword in first_col for keyword in ['설명', '입력', '예시', '주의']):
                    data_sections[current_section]['data'].append(row.tolist())
        
        return data_sections
        
    except Exception as e:
        print(f"지역 {region_code} 데이터 읽기 오류: {str(e)}")
        return {}

def create_integrated_data():
    """v3 생성기로 위임"""
    return v3_create_integrated_data()

if __name__ == "__main__":
    print("=== 지역별 데이터 통합하여 integrated_input_data.xlsx 생성 ===")
    create_integrated_data() 