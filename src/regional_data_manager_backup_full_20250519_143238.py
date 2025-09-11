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
REGION_TEMPLATES = {}
