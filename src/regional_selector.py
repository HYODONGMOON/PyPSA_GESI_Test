#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
지역 선택 모듈

한국 행정구역을 기반으로 에너지시스템 모델링을 위한 지역을 선택하는 모듈
"""

import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from matplotlib.collections import PatchCollection
import json
import math

# 기본 한국 행정구역 정보
KOREA_REGIONS = {
    'SEL': {
        'name': '서울특별시',
        'name_eng': 'Seoul',
        'center': (126.986, 37.566),
        'area': 605.21,  # km²
        'population': 9720846,  # 2021년 기준
        'color': 'lightcoral'
    },
    'BSN': {
        'name': '부산광역시',
        'name_eng': 'Busan',
        'center': (129.075, 35.180),
        'area': 770.04,
        'population': 3385601,
        'color': 'lightskyblue'
    },
    'DGU': {
        'name': '대구광역시',
        'name_eng': 'Daegu',
        'center': (128.601, 35.871),
        'area': 883.56,
        'population': 2418346,
        'color': 'lightgreen'
    },
    'ICN': {
        'name': '인천광역시',
        'name_eng': 'Incheon',
        'center': (126.705, 37.456),
        'area': 1063.27,
        'population': 2947217,
        'color': 'yellow'
    },
    'GWJ': {
        'name': '광주광역시',
        'name_eng': 'Gwangju',
        'center': (126.851, 35.160),
        'area': 501.24,
        'population': 1441611,
        'color': 'pink'
    },
    'DJN': {
        'name': '대전광역시',
        'name_eng': 'Daejeon',
        'center': (127.385, 36.351),
        'area': 539.84,
        'population': 1463882,
        'color': 'lightseagreen'
    },
    'USN': {
        'name': '울산광역시',
        'name_eng': 'Ulsan',
        'center': (129.311, 35.539),
        'area': 1062.04,
        'population': 1136017,
        'color': 'plum'
    },
    'SJN': {
        'name': '세종특별자치시',
        'name_eng': 'Sejong',
        'center': (127.289, 36.480),
        'area': 465.23,
        'population': 365309,
        'color': 'orange'
    },
    'GGD': {
        'name': '경기도',
        'name_eng': 'Gyeonggi',
        'center': (127.013, 37.275),
        'area': 10191.79,
        'population': 13530519,
        'color': 'lightsalmon'
    },
    'GWD': {
        'name': '강원도',
        'name_eng': 'Gangwon',
        'center': (128.318, 37.883),
        'area': 16826.37,
        'population': 1538492,
        'color': 'lightcyan'
    },
    'CBD': {
        'name': '충청북도',
        'name_eng': 'Chungbuk',
        'center': (127.705, 36.801),
        'area': 7407.30,
        'population': 1600957,
        'color': 'wheat'
    },
    'CND': {
        'name': '충청남도',
        'name_eng': 'Chungnam',
        'center': (126.799, 36.658),
        'area': 8245.41,
        'population': 2118670,
        'color': 'lightsteelblue'
    },
    'JBD': {
        'name': '전라북도',
        'name_eng': 'Jeonbuk',
        'center': (127.144, 35.820),
        'area': 8069.07,
        'population': 1792712,
        'color': 'mediumaquamarine'
    },
    'JND': {
        'name': '전라남도',
        'name_eng': 'Jeonnam',
        'center': (126.991, 34.868),
        'area': 12344.06,
        'population': 1851549,
        'color': 'thistle'
    },
    'GBD': {
        'name': '경상북도',
        'name_eng': 'Gyeongbuk',
        'center': (128.744, 36.566),
        'area': 19030.80,
        'population': 2641823,
        'color': 'paleturquoise'
    },
    'GND': {
        'name': '경상남도',
        'name_eng': 'Gyeongnam',
        'center': (128.241, 35.462),
        'area': 10540.29,
        'population': 3340216,
        'color': 'peachpuff'
    },
    'JJD': {
        'name': '제주특별자치도',
        'name_eng': 'Jeju',
        'center': (126.542, 33.387),
        'area': 1849.16,
        'population': 674635,
        'color': 'mediumpurple'
    },
}

class RegionalSelector:
    """지역 선택 및 관리 클래스"""
    
    def __init__(self, boundary_file=None):
        """초기화 함수
        
        Args:
            boundary_file (str, optional): 행정구역 경계 파일 경로
        """
        self.regions = KOREA_REGIONS
        self.selected_regions = []
        self.boundaries = None
        
        # 경계 파일이 제공된 경우 로드
        if boundary_file and os.path.exists(boundary_file):
            self.load_boundaries(boundary_file)
    
    def load_boundaries(self, file_path):
        """행정구역 경계 파일 로드
        
        Args:
            file_path (str): 경계 파일 경로 (GeoJSON 또는 JSON 포맷)
            
        Returns:
            bool: 로드 성공 여부
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                self.boundaries = json.load(f)
            return True
        except Exception as e:
            print(f"경계 파일 로드 오류: {str(e)}")
            return False
    
    def select_region(self, region_code):
        """지역 선택
        
        Args:
            region_code (str): 지역 코드 (예: 'SEL', 'GGD' 등)
            
        Returns:
            bool: 선택 성공 여부
        """
        if region_code in self.regions and region_code not in self.selected_regions:
            self.selected_regions.append(region_code)
            print(f"지역 선택: {self.regions[region_code]['name']}")
            return True
        return False
    
    def deselect_region(self, region_code):
        """지역 선택 해제
        
        Args:
            region_code (str): 지역 코드
            
        Returns:
            bool: 해제 성공 여부
        """
        if region_code in self.selected_regions:
            self.selected_regions.remove(region_code)
            print(f"지역 선택 해제: {self.regions[region_code]['name']}")
            return True
        return False
    
    def get_selected_regions(self):
        """선택된 지역 목록 반환
        
        Returns:
            list: 선택된 지역 코드 목록
        """
        return self.selected_regions
    
    def get_region_info(self, region_code):
        """지역 정보 반환
        
        Args:
            region_code (str): 지역 코드
            
        Returns:
            dict: 지역 정보
        """
        if region_code in self.regions:
            return self.regions[region_code]
        return None
    
    def draw_korea_map(self, highlight_selected=True, ax=None, figsize=(10, 12)):
        """한국 지도 그리기
        
        Args:
            highlight_selected (bool): 선택된 지역 강조 표시 여부
            ax (matplotlib.axes.Axes, optional): 그림을 그릴 축 객체
            figsize (tuple): 그림 크기
            
        Returns:
            tuple: (fig, ax) 그림 객체와 축 객체
        """
        if ax is None:
            fig, ax = plt.subplots(figsize=figsize)
        else:
            fig = ax.figure
        
        # 경계 파일이 있는 경우 정확한 경계 그리기
        if self.boundaries:
            # 경계 파일 기반 그리기 로직
            pass
        else:
            # 간단한 점 기반 표시
            for code, region in self.regions.items():
                color = region['color']
                alpha = 0.7
                
                # 선택된 지역 강조
                if highlight_selected and code in self.selected_regions:
                    alpha = 1.0
                    edgecolor = 'red'
                    linewidth = 2
                else:
                    edgecolor = 'black'
                    linewidth = 0.5
                
                # 원으로 표시 (실제 면적에 비례)
                area_scaled = math.sqrt(region['area']) / 10
                ax.add_patch(plt.Circle(
                    region['center'], 
                    radius=area_scaled, 
                    alpha=alpha,
                    color=color,
                    edgecolor=edgecolor,
                    linewidth=linewidth
                ))
                
                # 지역명 표시
                ax.annotate(region['name'], 
                           xy=region['center'],
                           ha='center', va='center',
                           fontsize=8)
        
        # 축 설정
        ax.set_aspect('equal')
        ax.set_title('대한민국 행정구역')
        
        # 위도/경도 범위 설정
        ax.set_xlim(125.5, 129.5)  # 대한민국 경도 범위
        ax.set_ylim(33.0, 38.5)    # 대한민국 위도 범위
        
        # 격자 및 레이블 제거
        ax.grid(False)
        ax.set_xticks([])
        ax.set_yticks([])
        
        return fig, ax
    
    def calculate_distance(self, region_code1, region_code2):
        """두 지역 간 거리 계산 (하버사인 공식)
        
        Args:
            region_code1 (str): 첫 번째 지역 코드
            region_code2 (str): 두 번째 지역 코드
            
        Returns:
            float: 거리 (km)
        """
        if region_code1 not in self.regions or region_code2 not in self.regions:
            return None
        
        # 두 지역의 중심 좌표
        lon1, lat1 = self.regions[region_code1]['center']
        lon2, lat2 = self.regions[region_code2]['center']
        
        # 하버사인 공식
        R = 6371.0  # 지구 반지름 (km)
        
        lat1_rad = math.radians(lat1)
        lon1_rad = math.radians(lon1)
        lat2_rad = math.radians(lat2)
        lon2_rad = math.radians(lon2)
        
        dlon = lon2_rad - lon1_rad
        dlat = lat2_rad - lat1_rad
        
        a = math.sin(dlat / 2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon / 2)**2
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
        
        distance = R * c
        return distance
    
    def get_all_distances(self):
        """선택된 모든 지역 쌍 간의 거리 계산
        
        Returns:
            dict: 지역 쌍과 거리 매핑
        """
        distances = {}
        
        for i, region1 in enumerate(self.selected_regions):
            for region2 in self.selected_regions[i+1:]:
                distance = self.calculate_distance(region1, region2)
                if distance is not None:
                    key = (region1, region2)
                    distances[key] = distance
        
        return distances
    
    def get_region_prefix(self, region_code):
        """지역 코드에서 파일명 접두사 생성
        
        Args:
            region_code (str): 지역 코드
            
        Returns:
            str: 접두사
        """
        if region_code in self.regions:
            return f"{region_code}_"
        return None

# 테스트 코드
if __name__ == "__main__":
    selector = RegionalSelector()
    
    # 지역 선택
    selector.select_region('SEL')  # 서울
    selector.select_region('JBD')  # 전북
    
    # 지도 그리기
    fig, ax = selector.draw_korea_map()
    plt.show()
    
    # 거리 계산
    distances = selector.get_all_distances()
    for (r1, r2), dist in distances.items():
        name1 = selector.get_region_info(r1)['name']
        name2 = selector.get_region_info(r2)['name']
        print(f"{name1} - {name2} 간 거리: {dist:.1f} km") 