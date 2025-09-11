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
REGION_TEMPLATES = {}
class RegionalDataManager:
    """지역별 에너지시스템 데이터 관리 클래스"""
    def initialize_region(self, region_code):
        """지역 초기화
        
        Args:
            region_code (str): 지역 코드
        """
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
        """구성요소 추가
        
        Args:
            region_code (str): 지역 코드
            component_type (str): 구성요소 유형 (buses, generators, loads, lines, stores, links)
            data (dict): 구성요소 데이터
        """
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
        """지역간 연결 추가
        
        Args:
            region1 (str): 시작 지역 코드
            region2 (str): 도착 지역 코드
            connection_data (dict, optional): 연결 속성 데이터
        """
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
        """모든 지역 데이터 병합
        
        Returns:
            dict: 구성요소별로 병합된 데이터프레임
        """
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
        """병합된 데이터를 엑셀 파일로 내보내기
        
        Args:
            output_path (str): 출력 파일 경로
        """
        merged_data = self.merge_data()
        
        with pd.ExcelWriter(output_path) as writer:
            for component, df in merged_data.items():
                if not df.empty:
                    df.to_excel(writer, sheet_name=component, index=False)
                    
        print(f"통합 데이터가 '{output_path}'에 저장되었습니다.")
    
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
