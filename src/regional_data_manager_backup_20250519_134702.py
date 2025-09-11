#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
지역별 데이터 관리 모듈

지역별 에너지시스템 데이터를 관리하는 모듈
"""

import os
import pandas as pd
import numpy as np
from datetime import datetime
import json
import traceback

# 지역 선택기 임포트
from regional_selector import RegionalSelector

# 구성요소 템플릿 정의
COMPONENT_TEMPLATES = {
    'buses': {
        'columns': ['name', 'region', 'v_nom', 'carrier', 'x', 'y'],
        'required': ['name', 'v_nom', 'carrier'],
        'types': {
            'name': str,
            'region': str,
            'v_nom': float,
            'carrier': str,
            'x': float,
            'y': float
        }
    },
    'generators': {
        'columns': ['name', 'region', 'bus', 'carrier', 'p_nom', 'p_nom_extendable',
                  'p_nom_min', 'p_nom_max', 'marginal_cost', 'capital_cost',
                  'efficiency', 'p_max_pu', 'committable'],
        'required': ['name', 'bus', 'carrier', 'p_nom'],
        'types': {
            'name': str,
            'region': str,
            'bus': str,
            'carrier': str,
            'p_nom': float,
            'p_nom_extendable': bool,
            'p_nom_min': float,
            'p_nom_max': float,
            'marginal_cost': float,
            'capital_cost': float,
            'efficiency': float,
            'p_max_pu': float,
            'committable': bool
        }
    },
    'loads': {
        'columns': ['name', 'region', 'bus', 'carrier', 'p_set'],
        'required': ['name', 'bus', 'p_set'],
        'types': {
            'name': str,
            'region': str,
            'bus': str,
            'carrier': str,
            'p_set': float
        }
    },
    'lines': {
        'columns': ['name', 'region', 'bus0', 'bus1', 'carrier', 'x', 'r', 's_nom',
                   's_nom_extendable', 's_nom_min', 's_nom_max', 'length', 'capital_cost'],
        'required': ['name', 'bus0', 'bus1', 's_nom'],
        'types': {
            'name': str,
            'region': str,
            'bus0': str,
            'bus1': str,
            'carrier': str,
            'x': float,
            'r': float,
            's_nom': float,
            's_nom_extendable': bool,
            's_nom_min': float,
            's_nom_max': float,
            'length': float,
            'capital_cost': float
        }
    },
    'stores': {
        'columns': ['name', 'region', 'bus', 'carrier', 'e_nom', 'e_nom_extendable',
                   'e_cyclic', 'standing_loss', 'efficiency_store', 'efficiency_dispatch',
                   'e_initial', 'e_nom_max', 'e_nom_min', 'capital_cost'],
        'required': ['name', 'bus', 'e_nom'],
        'types': {
            'name': str,
            'region': str,
            'bus': str,
            'carrier': str,
            'e_nom': float,
            'e_nom_extendable': bool,
            'e_cyclic': bool,
            'standing_loss': float,
            'efficiency_store': float,
            'efficiency_dispatch': float,
            'e_initial': float,
            'e_nom_max': float,
            'e_nom_min': float,
            'capital_cost': float
        }
    },
    'links': {
        'columns': ['name', 'region', 'bus0', 'bus1', 'bus2', 'bus3', 'efficiency0',
                   'efficiency1', 'efficiency2', 'efficiency3', 'p_nom', 'p_nom_extendable',
                   'p_nom_min', 'p_nom_max', 'capital_cost', 'marginal_cost'],
        'required': ['name', 'bus0', 'p_nom'],
        'types': {
            'name': str,
            'region': str,
            'bus0': str,
            'bus1': str,
            'bus2': str,
            'bus3': str,
            'efficiency0': float,
            'efficiency1': float,
            'efficiency2': float,
            'efficiency3': float,
            'p_nom': float,
            'p_nom_extendable': bool,
            'p_nom_min': float,
            'p_nom_max': float,
            'capital_cost': float,
            'marginal_cost': float
        }
    },
}

# 지역별 기본 데이터 템플릿
REGION_TEMPLATES = {
    'SEL': {  # 서울
        'buses': [
            {'name': 'Main_EL', 'v_nom': 345.0, 'carrier': 'AC', 'x': 126.986, 'y': 37.566},
            {'name': 'Hydrogen', 'v_nom': 10.0, 'carrier': 'hydrogen', 'x': 126.980, 'y': 37.560}
        ],
        'generators': [
            {'name': 'PV1', 'bus': 'Main_EL', 'carrier': 'solar', 'p_nom': 1000.0, 'p_nom_extendable': True, 
             'p_nom_max': 2000.0, 'marginal_cost': 0.0, 'capital_cost': 800000.0},
            {'name': 'Nuclear1', 'bus': 'Main_EL', 'carrier': 'nuclear', 'p_nom': 5000.0, 'p_nom_extendable': False,
             'marginal_cost': 10.0, 'capital_cost': 5000000.0}
        ],
        'loads': [
            {'name': 'Demand1', 'bus': 'Main_EL', 'p_set': 15000.0},
            {'name': 'H2_Demand', 'bus': 'Hydrogen', 'p_set': 1000.0}
        ],
        'stores': [
            {'name': 'Battery', 'bus': 'Main_EL', 'carrier': 'electricity', 'e_nom': 10000.0, 'e_nom_extendable': True,
             'e_cyclic': True, 'efficiency_store': 0.95, 'efficiency_dispatch': 0.95}
        ]
    },
    'JBD': {  # 전라북도
        'buses': [
            {'name': 'Main_EL', 'v_nom': 345.0, 'carrier': 'AC', 'x': 127.144, 'y': 35.820},
            {'name': 'Hydrogen', 'v_nom': 10.0, 'carrier': 'hydrogen', 'x': 127.140, 'y': 35.815}
        ],
        'generators': [
            {'name': 'PV1', 'bus': 'Main_EL', 'carrier': 'solar', 'p_nom': 3000.0, 'p_nom_extendable': True, 
             'p_nom_max': 5000.0, 'marginal_cost': 0.0, 'capital_cost': 700000.0},
            {'name': 'WT1', 'bus': 'Main_EL', 'carrier': 'wind', 'p_nom': 2000.0, 'p_nom_extendable': True,
             'p_nom_max': 4000.0, 'marginal_cost': 0.0, 'capital_cost': 1000000.0}
        ],
        'loads': [
            {'name': 'Demand1', 'bus': 'Main_EL', 'p_set': 5000.0},
            {'name': 'H2_Demand', 'bus': 'Hydrogen', 'p_set': 500.0}
        ],
        'stores': [
            {'name': 'Battery', 'bus': 'Main_EL', 'carrier': 'electricity', 'e_nom': 5000.0, 'e_nom_extendable': True,
             'e_cyclic': True, 'efficiency_store': 0.95, 'efficiency_dispatch': 0.95}
        ]
    }
}

class RegionalDataManager:
    """지역별 에너지시스템 데이터 관리 클래스"""
    
    def __init__(self, region_selector):
        """초기화 함수
        
        Args:
            region_selector (RegionalSelector): 지역 선택 객체
        """
        self.region_selector = region_selector
        self.regional_data = {}  # 지역별 데이터 저장
        self.connections = []    # 지역간 연결 저장
        
        # 기본 템플릿 로드
        self.templates = COMPONENT_TEMPLATES
        self.region_templates = REGION_TEMPLATES
    
    def initialize_region(self, region_code):
        """특정 지역의 기본 데이터 초기화
        
        Args:
            region_code (str): 지역 코드
            
        Returns:
            bool: 초기화 성공 여부
        """
        if region_code not in self.region_selector.regions:
            print(f"오류: 유효하지 않은 지역 코드 '{region_code}'")
            return False
        
        # 이미 초기화된 지역인 경우 데이터 초기화
        if region_code in self.regional_data:
            print(f"지역 '{region_code}'를 다시 초기화합니다.")
            # 기존 데이터 삭제
            del self.regional_data[region_code]
        
        # 빈 데이터프레임으로 초기화
        self.regional_data[region_code] = {}
        for component, template in self.templates.items():
            self.regional_data[region_code][component] = pd.DataFrame(columns=template['columns'])
        
        # 지역 정보 추가
        region_info = self.region_selector.get_region_info(region_code)
        self.regional_data[region_code]['info'] = region_info
        
        # 기본 템플릿이 있는 경우 적용
        if region_code in self.region_templates:
            self._apply_template(region_code)
        
        print(f"지역 '{region_info['name']}' 초기화 완료")
        return True
    
    def _apply_template(self, region_code):
        """지역에 기본 템플릿 적용
        
        Args:
            region_code (str): 지역 코드
        """
        if region_code not in self.region_templates:
            return
        
        template = self.region_templates[region_code]
        prefix = self.region_selector.get_region_prefix(region_code)
        
        # 템플릿의 각 구성요소 적용
        for component, items in template.items():
            if component not in self.regional_data[region_code]:
                continue
                
            # 각 항목에 지역 정보 추가 및 이름 접두사 적용
            for item in items:
                item_copy = item.copy()
                
                # 지역 정보 추가
                item_copy['region'] = region_code
                
                # 이름에 접두사 추가
                if 'name' in item_copy:
                    item_copy['name'] = f"{prefix}{item_copy['name']}"
                
                # 버스 참조에 접두사 추가
                if 'bus' in item_copy:
                    item_copy['bus'] = f"{prefix}{item_copy['bus']}"
                if 'bus0' in item_copy:
                    item_copy['bus0'] = f"{prefix}{item_copy['bus0']}"
                if 'bus1' in item_copy and item_copy['bus1']:
                    item_copy['bus1'] = f"{prefix}{item_copy['bus1']}"
                
                # 데이터프레임에 추가
                self.add_component(region_code, component, item_copy)
    
    def add_component(self, region_code, component_type, data):
        """구성요소 추가
        
        Args:
            region_code (str): 지역 코드
            component_type (str): 구성요소 유형 (buses, generators 등)
            data (dict): 구성요소 데이터
            
        Returns:
            bool: 추가 성공 여부
        """
        # 지역 초기화 확인
        if region_code not in self.regional_data:
            print(f"오류: 지역 '{region_code}'가 초기화되지 않았습니다.")
            return False
        
        # 구성요소 유형 확인
        if component_type not in self.templates:
            print(f"오류: 유효하지 않은 구성요소 유형 '{component_type}'")
            return False
        
        try:
            # 필수 필드 확인
            for field in self.templates[component_type]['required']:
                if field not in data:
                    print(f"오류: 필수 필드 '{field}'가 누락되었습니다.")
                    return False
            
            # 데이터 유형 변환
            for field, field_type in self.templates[component_type]['types'].items():
                if field in data and data[field] is not None:
                    try:
                        data[field] = field_type(data[field])
                    except ValueError:
                        print(f"경고: 필드 '{field}'의 값 '{data[field]}'을(를) {field_type.__name__} 유형으로 변환할 수 없습니다.")
            
            # 빈 필드 제거
            data = {k: v for k, v in data.items() if v is not None}
            
            # 데이터프레임에 추가
            df = self.regional_data[region_code][component_type]
            df_new = pd.DataFrame([data])
            self.regional_data[region_code][component_type] = pd.concat([df, df_new], ignore_index=True)
            
            return True
            
        except Exception as e:
            print(f"구성요소 추가 중 오류 발생: {str(e)}")
            traceback.print_exc()
            return False
    
    def add_connection(self, region_code1, region_code2, connection_data=None):
        """두 지역 간 연결 추가
        
        Args:
            region_code1 (str): 첫 번째 지역 코드
            region_code2 (str): 두 번째 지역 코드
            connection_data (dict, optional): 연결 정보. 없으면 기본값 사용
            
        Returns:
            bool: 추가 성공 여부
        """
        # 지역 초기화 확인
        if region_code1 not in self.regional_data or region_code2 not in self.regional_data:
            print(f"오류: 두 지역 모두 초기화되어야 합니다.")
            return False
        
        try:
            # 기본 연결 정보 생성
            prefix1 = self.region_selector.get_region_prefix(region_code1)
            prefix2 = self.region_selector.get_region_prefix(region_code2)
            
            distance = self.region_selector.calculate_distance(region_code1, region_code2)
            
            # 기본 연결 정보
            default_connection = {
                'name': f"{region_code1}_to_{region_code2}_Line",
                'bus0': f"{prefix1}Main_EL",
                'bus1': f"{prefix2}Main_EL",
                'carrier': 'AC',
                'x': distance * 0.0004,  # 거리에 비례한 리액턴스
                'r': distance * 0.0001,  # 거리에 비례한 저항
                's_nom': 1000.0,  # 기본 용량 (MW)
                'length': distance,
                'v_nom': 345.0  # 기본 전압 (kV)
            }
            
            # 사용자 제공 데이터로 덮어쓰기
            if connection_data:
                default_connection.update(connection_data)
            
            # 연결 추가
            self.connections.append(default_connection)
            
            print(f"지역 '{region_code1}'과(와) '{region_code2}' 간 연결 추가됨 (거리: {distance:.1f} km)")
            return True
            
        except Exception as e:
            print(f"연결 추가 중 오류 발생: {str(e)}")
            traceback.print_exc()
            return False
    
    def import_regional_data(self, region_code, file_path):
        """외부 파일에서 지역 데이터 가져오기
        
        Args:
            region_code (str): 지역 코드
            file_path (str): 데이터 파일 경로 (Excel)
            
        Returns:
            bool: 가져오기 성공 여부
        """
        if not os.path.exists(file_path):
            print(f"오류: 파일 '{file_path}'를 찾을 수 없습니다.")
            return False
        
        try:
            # 지역 초기화
            if region_code not in self.regional_data:
                self.initialize_region(region_code)
            
            # 엑셀 파일 읽기
            xls = pd.ExcelFile(file_path)
            
            # 각 시트 처리
            for sheet_name in xls.sheet_names:
                if sheet_name in self.templates:
                    # 데이터 읽기
                    df = pd.read_excel(file_path, sheet_name=sheet_name)
                    df.columns = df.columns.str.strip()
                    
                    # 각 행 처리
                    for _, row in df.iterrows():
                        data = row.to_dict()
                        data['region'] = region_code  # 지역 정보 추가
                        
                        # 이름에 접두사 추가
                        prefix = self.region_selector.get_region_prefix(region_code)
                        if 'name' in data and not str(data['name']).startswith(prefix):
                            data['name'] = f"{prefix}{data['name']}"
                        
                        # 버스 참조에 접두사 추가
                        if 'bus' in data and data['bus'] and not str(data['bus']).startswith(prefix):
                            data['bus'] = f"{prefix}{data['bus']}"
                        if 'bus0' in data and data['bus0'] and not str(data['bus0']).startswith(prefix):
                            data['bus0'] = f"{prefix}{data['bus0']}"
                        if 'bus1' in data and data['bus1'] and not str(data['bus1']).startswith(prefix):
                            data['bus1'] = f"{prefix}{data['bus1']}"
                        
                        # 구성요소 추가
                        self.add_component(region_code, sheet_name, data)
            
            print(f"파일 '{file_path}'에서 지역 '{region_code}' 데이터를 가져왔습니다.")
            return True
            
        except Exception as e:
            print(f"데이터 가져오기 중 오류 발생: {str(e)}")
            traceback.print_exc()
            return False
    
    def export_regional_data(self, region_code, file_path):
        """지역 데이터를 외부 파일로 내보내기
        
        Args:
            region_code (str): 지역 코드
            file_path (str): 저장할 파일 경로
            
        Returns:
            bool: 내보내기 성공 여부
        """
        if region_code not in self.regional_data:
            print(f"오류: 지역 '{region_code}'가 초기화되지 않았습니다.")
            return False
        
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # 각 구성요소 저장
                for component, df in self.regional_data[region_code].items():
                    if component != 'info' and not df.empty:
                        df.to_excel(writer, sheet_name=component, index=False)
            
            print(f"지역 '{region_code}' 데이터를 '{file_path}'에 저장했습니다.")
            return True
            
        except Exception as e:
            print(f"데이터 내보내기 중 오류 발생: {str(e)}")
            traceback.print_exc()
            return False
    
    def merge_data(self):
        """모든 지역 데이터를 통합
        
        Returns:
            dict: 통합된 데이터
        """
        if not self.regional_data:
            print("오류: 통합할 지역 데이터가 없습니다.")
            return None
        
        try:
            # 통합 데이터 초기화
            merged_data = {}
            for component in self.templates.keys():
                merged_data[component] = pd.DataFrame()
            
            # 각 지역 데이터 통합 (중복 제거)
            processed_regions = set()  # 이미 처리된 지역 추적
            
            for region_code, data in self.regional_data.items():
                # 이미 처리된 지역은 건너뛰기
                if region_code in processed_regions:
                    print(f"지역 '{region_code}'는 이미 처리되었습니다. 건너뜁니다.")
                    continue
                    
                processed_regions.add(region_code)
                print(f"지역 '{region_code}' 데이터 병합 중...")
                
                for component, df in data.items():
                    if component != 'info' and not df.empty:
                        # 이름 기반으로 중복 확인 및 제거
                        if not merged_data[component].empty and 'name' in df.columns:
                            # 이미 있는 이름 목록
                            existing_names = merged_data[component]['name'].tolist() if 'name' in merged_data[component].columns else []
                            
                            # 중복되지 않는 항목만 필터링
                            df_filtered = df[~df['name'].isin(existing_names)]
                            
                            if len(df_filtered) < len(df):
                                print(f"  - {component}: {len(df) - len(df_filtered)}개 중복 항목 제거됨")
                            
                            # 필터링된 데이터 병합
                            merged_data[component] = pd.concat([merged_data[component], df_filtered], ignore_index=True)
                        else:
                            # 기존 데이터가 없으면 그대로 추가
                            merged_data[component] = pd.concat([merged_data[component], df], ignore_index=True)
            
            # 지역간 연결 추가
            if self.connections:
                connections_df = pd.DataFrame(self.connections)
                
                if 'lines' in merged_data:
                    # 기존 lines 데이터프레임과 연결
                    merged_data['lines'] = pd.concat([merged_data['lines'], connections_df], ignore_index=True)
                else:
                    merged_data['lines'] = connections_df
            
            print("지역 데이터 통합 완료")
            return merged_data
            
        except Exception as e:
            print(f"데이터 통합 중 오류 발생: {str(e)}")
            traceback.print_exc()
            return None
    
    def export_merged_data(self, file_path):
        """통합된 데이터를 엑셀 파일로 내보내기
        
        Args:
            file_path (str): 저장할 파일 경로
            
        Returns:
            bool: 내보내기 성공 여부
        """
        merged_data = self.merge_data()
        if not merged_data:
            return False
        
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # 각 구성요소 저장
                for component, df in merged_data.items():
                    if not df.empty:
                        # 표준 입력 형식에 맞게 변환
                        if component == 'buses':
                            # 지역 필드 제거
                            if 'region' in df.columns:
                                df = df.drop(columns=['region'])
                        
                        df.to_excel(writer, sheet_name=component, index=False)
                
                # 시간 설정 시트 추가
                self._add_timeseries_sheet(writer)
                
                # 재생에너지 패턴 시트 추가
                self._add_patterns_sheet(writer)
            
            print(f"통합 데이터를 '{file_path}'에 저장했습니다.")
            return True
            
        except Exception as e:
            print(f"통합 데이터 내보내기 중 오류 발생: {str(e)}")
            traceback.print_exc()
            return False
    
    def _add_timeseries_sheet(self, writer):
        """시간 설정 시트 추가
        
        Args:
            writer (ExcelWriter): 엑셀 작성자 객체
        """
        # 기본 시간 설정
        timeseries = pd.DataFrame({
            'start_time': [datetime(2023, 1, 1, 0, 0, 0)],
            'end_time': [datetime(2024, 1, 1, 0, 0, 0)],
            'frequency': ['1h']
        })
        
        timeseries.to_excel(writer, sheet_name='timeseries', index=False)
    
    def _add_patterns_sheet(self, writer):
        """재생에너지 패턴 시트 추가
        
        Args:
            writer (ExcelWriter): 엑셀 작성자 객체
        """
        # 기본 패턴 (간단한 예시)
        hours = list(range(1, 8761))  # 1년(8760시간)
        
        # 태양광 패턴 (낮에 높고 밤에 0)
        pv_pattern = []
        for h in range(1, 8761):
            hour_of_day = (h - 1) % 24
            if 6 <= hour_of_day < 18:  # 낮 시간
                value = 0.5 + 0.5 * np.sin(np.pi * (hour_of_day - 6) / 12)
            else:  # 밤 시간
                value = 0.0
            pv_pattern.append(value)
        
        # 풍력 패턴 (더 랜덤하게)
        np.random.seed(42)  # 재현성을 위한 시드 설정
        wt_pattern = 0.2 + 0.6 * np.random.rand(8760)
        
        # 데이터프레임 생성
        patterns = pd.DataFrame({
            'hour': hours,
            'PV': pv_pattern,
            'WT': wt_pattern
        })
        
        patterns.to_excel(writer, sheet_name='renewable_patterns', index=False)

# 테스트 코드
if __name__ == "__main__":
    # 지역 선택기 생성
    selector = RegionalSelector()
    
    # 데이터 관리자 생성
    manager = RegionalDataManager(selector)
    
    # 지역 선택 및 초기화
    selector.select_region('SEL')  # 서울
    selector.select_region('JBD')  # 전북
    
    manager.initialize_region('SEL')
    manager.initialize_region('JBD')
    
    # 지역간 연결 추가
    manager.add_connection('SEL', 'JBD')
    
    # 통합 데이터 내보내기
    manager.export_merged_data('integrated_system.xlsx') 