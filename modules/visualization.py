#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
시각화 모듈

최적화 결과를 다양한 그래프와 차트로 시각화합니다.
"""

import os
import logging
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from datetime import datetime

logger = logging.getLogger("PyPSA-HD.Visualization")

# 그래프에 한글 폰트 설정
import platform
if platform.system() == "Windows":
    plt.rc('font', family='Malgun Gothic')  # Windows의 경우 맑은 고딕
elif platform.system() == "Darwin":
    plt.rc('font', family='AppleGothic')   # Mac의 경우 애플고딕
plt.rcParams['axes.unicode_minus'] = False  # 마이너스 기호 깨짐 방지

class Visualizer:
    """결과 시각화 클래스"""
    
    def __init__(self, config):
        """초기화 함수
        
        Args:
            config (dict): 설정 정보
        """
        self.config = config
        self.output_dir = config.get('output_dir', 'results')
        self.plot_dir = os.path.join(self.output_dir, 'plots')
        os.makedirs(self.plot_dir, exist_ok=True)
        
        # 시각화 설정
        self.dpi = config.get('visualization', {}).get('dpi', 300)
        self.figsize = config.get('visualization', {}).get('figsize', (10, 6))
        self.map_enabled = config.get('visualization', {}).get('map_enabled', True)
    
    def visualize_results(self, network, input_data):
        """최적화 결과 시각화
        
        Args:
            network (pypsa.Network): 최적화된 PyPSA 네트워크 객체
            input_data (dict): 입력 데이터
            
        Returns:
            bool: 시각화 성공 여부
        """
        try:
            if not hasattr(network, 'objective') or network.objective is None:
                logger.warning("최적화 결과가 없어 시각화할 수 없습니다.")
                return False
            
            logger.info("결과 시각화 시작...")
            
            # 시간대별 발전 및 부하 시각화
            self.plot_generation_profile(network)
            
            # 발전원별 발전량 시각화
            self.plot_generation_by_carrier(network)
            
            # 저장장치 상태 시각화
            if not network.stores.empty:
                self.plot_storage_state(network)
            
            # 선로 및 링크 이용률 시각화
            self.plot_line_loading(network)
            self.plot_link_loading(network)
            
            # 버스별 한계가격 시각화
            self.plot_marginal_prices(network)
            
            # 네트워크 맵 시각화 (설정된 경우)
            if self.map_enabled:
                self.plot_network_map(network, input_data)
            
            logger.info("결과 시각화 완료")
            return True
            
        except Exception as e:
            logger.error(f"시각화 중 오류 발생: {str(e)}")
            return False
    
    def plot_generation_profile(self, network):
        """시간대별 발전 프로필 그래프
        
        Args:
            network (pypsa.Network): 최적화된 PyPSA 네트워크 객체
        """
        if network.generators.empty or network.generators_t.p.empty:
            logger.warning("발전기 출력 데이터가 없어 발전 프로필을 시각화할 수 없습니다.")
            return
        
        try:
            # 발전원별 그룹화
            gen_by_carrier = network.generators.groupby('carrier')
            
            # 타임스텝과 캐리어별 발전량 계산
            generation = pd.DataFrame(index=network.snapshots)
            
            for carrier, gens in gen_by_carrier.groups.items():
                generation[carrier] = network.generators_t.p[gens].sum(axis=1)
            
            # 부하 데이터 추가
            if not network.loads_t.p.empty:
                generation['load'] = network.loads_t.p.sum(axis=1)
            
            # 그래프 그리기
            plt.figure(figsize=self.figsize)
            
            # 발전원별 그래프
            ax = generation.drop('load', axis=1, errors='ignore').plot.area(
                figsize=self.figsize,
                linewidth=0,
                xlabel='시간',
                ylabel='발전량 (MW)'
            )
            
            # 부하 곡선 추가
            if 'load' in generation.columns:
                generation['load'].plot(
                    ax=ax,
                    color='black',
                    linestyle='-',
                    linewidth=2,
                    label='총 부하'
                )
            
            plt.title('시간대별 발전량 및 부하')
            plt.grid(True, alpha=0.3)
            plt.legend(loc='upper right')
            
            # 그래프 저장
            filename = os.path.join(self.plot_dir, 'generation_profile.png')
            plt.savefig(filename, dpi=self.dpi, bbox_inches='tight')
            plt.close()
            
            logger.info(f"발전 프로필 그래프를 '{filename}'에 저장했습니다.")
            
        except Exception as e:
            logger.error(f"발전 프로필 그래프 생성 중 오류 발생: {str(e)}")
    
    def plot_generation_by_carrier(self, network):
        """발전원별 총 발전량 그래프
        
        Args:
            network (pypsa.Network): 최적화된 PyPSA 네트워크 객체
        """
        if network.generators.empty or network.generators_t.p.empty:
            logger.warning("발전기 출력 데이터가 없어 발전원별 그래프를 시각화할 수 없습니다.")
            return
        
        try:
            # 발전원별 총 발전량 계산
            gen_by_carrier = network.generators.groupby('carrier')
            total_by_carrier = pd.Series(0.0, index=gen_by_carrier.groups.keys())
            
            for carrier, gens in gen_by_carrier.groups.items():
                total_by_carrier[carrier] = network.generators_t.p[gens].sum().sum()
            
            # 그래프 그리기
            plt.figure(figsize=self.figsize)
            
            # 색상 매핑
            colors = {
                'solar': 'gold',
                'wind': 'skyblue',
                'nuclear': 'purple',
                'coal': 'brown',
                'gas': 'gray',
                'electricity': 'blue'
            }
            
            # 발전원별 색상 적용
            carrier_colors = [colors.get(carrier, 'orange') for carrier in total_by_carrier.index]
            
            # 파이 차트
            plt.pie(
                total_by_carrier.values,
                labels=total_by_carrier.index,
                autopct='%1.1f%%',
                startangle=90,
                colors=carrier_colors
            )
            
            plt.title('발전원별 총 발전량')
            plt.axis('equal')  # 원형으로 유지
            
            # 그래프 저장
            filename = os.path.join(self.plot_dir, 'generation_by_carrier.png')
            plt.savefig(filename, dpi=self.dpi, bbox_inches='tight')
            plt.close()
            
            logger.info(f"발전원별 발전량 그래프를 '{filename}'에 저장했습니다.")
            
        except Exception as e:
            logger.error(f"발전원별 그래프 생성 중 오류 발생: {str(e)}")
    
    def plot_storage_state(self, network):
        """저장장치 상태 그래프
        
        Args:
            network (pypsa.Network): 최적화된 PyPSA 네트워크 객체
        """
        if network.stores.empty or network.stores_t.e.empty:
            logger.warning("저장장치 데이터가 없어 저장장치 상태를 시각화할 수 없습니다.")
            return
        
        try:
            # 모든 저장장치의 에너지 상태
            storage_energy = network.stores_t.e
            
            # 그래프 그리기
            plt.figure(figsize=self.figsize)
            
            ax = storage_energy.plot(
                figsize=self.figsize,
                linewidth=2,
                xlabel='시간',
                ylabel='저장 에너지 (MWh)'
            )
            
            plt.title('저장장치 에너지 상태')
            plt.grid(True, alpha=0.3)
            plt.legend(loc='best')
            
            # 그래프 저장
            filename = os.path.join(self.plot_dir, 'storage_state.png')
            plt.savefig(filename, dpi=self.dpi, bbox_inches='tight')
            plt.close()
            
            # 저장장치 충방전 그래프
            if not network.stores_t.p.empty:
                storage_power = network.stores_t.p
                
                plt.figure(figsize=self.figsize)
                
                ax = storage_power.plot(
                    figsize=self.figsize,
                    linewidth=2,
                    xlabel='시간',
                    ylabel='충방전 전력 (MW)'
                )
                
                plt.title('저장장치 충방전 전력')
                plt.grid(True, alpha=0.3)
                plt.axhline(y=0, color='black', linestyle='-', alpha=0.3)
                plt.legend(loc='best')
                
                # 그래프 저장
                filename = os.path.join(self.plot_dir, 'storage_power.png')
                plt.savefig(filename, dpi=self.dpi, bbox_inches='tight')
                plt.close()
            
            logger.info(f"저장장치 상태 그래프를 저장했습니다.")
            
        except Exception as e:
            logger.error(f"저장장치 그래프 생성 중 오류 발생: {str(e)}")
    
    def plot_line_loading(self, network):
        """선로 이용률 그래프
        
        Args:
            network (pypsa.Network): 최적화된 PyPSA 네트워크 객체
        """
        if network.lines.empty or network.lines_t.p0.empty:
            logger.warning("선로 데이터가 없어 선로 이용률을 시각화할 수 없습니다.")
            return
        
        try:
            # 선로 이용률 계산
            loading = network.lines_t.p0.abs() / network.lines.s_nom.values * 100
            
            # 선로별 최대 이용률
            max_loading = loading.max()
            
            # 그래프 그리기 (상위 10개 선로)
            plt.figure(figsize=self.figsize)
            
            ax = max_loading.nlargest(10).plot.bar(
                figsize=self.figsize,
                rot=45,
                ylabel='최대 이용률 (%)'
            )
            
            plt.title('주요 선로 최대 이용률 (상위 10개)')
            plt.grid(True, alpha=0.3)
            plt.axhline(y=100, color='red', linestyle='--', alpha=0.7, label='용량 한계')
            plt.tight_layout()
            
            # 그래프 저장
            filename = os.path.join(self.plot_dir, 'line_loading.png')
            plt.savefig(filename, dpi=self.dpi, bbox_inches='tight')
            plt.close()
            
            logger.info(f"선로 이용률 그래프를 '{filename}'에 저장했습니다.")
            
        except Exception as e:
            logger.error(f"선로 이용률 그래프 생성 중 오류 발생: {str(e)}")
    
    def plot_link_loading(self, network):
        """링크 이용률 그래프
        
        Args:
            network (pypsa.Network): 최적화된 PyPSA 네트워크 객체
        """
        if network.links.empty or network.links_t.p0.empty:
            logger.warning("링크 데이터가 없어 링크 이용률을 시각화할 수 없습니다.")
            return
        
        try:
            # 링크 이용률 계산
            loading = network.links_t.p0.abs() / network.links.p_nom.values * 100
            
            # 링크별 최대 이용률
            max_loading = loading.max()
            
            # 그래프 그리기 (상위 10개 링크)
            plt.figure(figsize=self.figsize)
            
            ax = max_loading.nlargest(10).plot.bar(
                figsize=self.figsize,
                rot=45,
                ylabel='최대 이용률 (%)'
            )
            
            plt.title('주요 링크 최대 이용률 (상위 10개)')
            plt.grid(True, alpha=0.3)
            plt.axhline(y=100, color='red', linestyle='--', alpha=0.7, label='용량 한계')
            plt.tight_layout()
            
            # 그래프 저장
            filename = os.path.join(self.plot_dir, 'link_loading.png')
            plt.savefig(filename, dpi=self.dpi, bbox_inches='tight')
            plt.close()
            
            logger.info(f"링크 이용률 그래프를 '{filename}'에 저장했습니다.")
            
        except Exception as e:
            logger.error(f"링크 이용률 그래프 생성 중 오류 발생: {str(e)}")
    
    def plot_marginal_prices(self, network):
        """버스별 한계가격 그래프
        
        Args:
            network (pypsa.Network): 최적화된 PyPSA 네트워크 객체
        """
        if network.buses_t.marginal_price.empty:
            logger.warning("한계가격 데이터가 없어 한계가격을 시각화할 수 없습니다.")
            return
        
        try:
            # 한계가격 평균 및 최대값
            mean_prices = network.buses_t.marginal_price.mean()
            max_prices = network.buses_t.marginal_price.max()
            
            # 그래프 그리기
            plt.figure(figsize=self.figsize)
            
            # 상위 10개 버스
            top_buses = mean_prices.nlargest(10)
            
            ax = top_buses.plot.bar(
                figsize=self.figsize,
                rot=45,
                ylabel='평균 한계가격 ($/MWh)'
            )
            
            plt.title('주요 버스 평균 한계가격 (상위 10개)')
            plt.grid(True, alpha=0.3)
            plt.tight_layout()
            
            # 그래프 저장
            filename = os.path.join(self.plot_dir, 'marginal_prices.png')
            plt.savefig(filename, dpi=self.dpi, bbox_inches='tight')
            plt.close()
            
            # 시간대별 평균 한계가격 그래프
            plt.figure(figsize=self.figsize)
            
            ax = network.buses_t.marginal_price.mean(axis=1).plot(
                figsize=self.figsize,
                linewidth=2,
                xlabel='시간',
                ylabel='평균 한계가격 ($/MWh)'
            )
            
            plt.title('시간대별 평균 한계가격')
            plt.grid(True, alpha=0.3)
            
            # 그래프 저장
            filename = os.path.join(self.plot_dir, 'marginal_prices_by_time.png')
            plt.savefig(filename, dpi=self.dpi, bbox_inches='tight')
            plt.close()
            
            logger.info(f"한계가격 그래프를 저장했습니다.")
            
        except Exception as e:
            logger.error(f"한계가격 그래프 생성 중 오류 발생: {str(e)}")
    
    def plot_network_map(self, network, input_data):
        """네트워크 맵 시각화
        
        Args:
            network (pypsa.Network): 최적화된 PyPSA 네트워크 객체
            input_data (dict): 입력 데이터
        """
        try:
            # 외부 지도 시각화 모듈 로드 (korea_map.py)
            try:
                from korea_map import KoreaMapVisualizer
                visualizer = KoreaMapVisualizer()
                
                # 지도 데이터 로드
                if visualizer.load_map_data():
                    # 지도 시각화
                    map_file = os.path.join(self.plot_dir, 'network_map.png')
                    visualizer.plot_korea_map(save_path=map_file)
                    
                    # 연결 정보 시각화 (GIS 시트가 있는 경우)
                    if 'GIS' in input_data:
                        connections_file = os.path.join(self.plot_dir, 'network_with_connections.png')
                        visualizer.add_transmission_lines(input_data, 'GIS', save_path=connections_file)
                    
                    logger.info(f"네트워크 맵을 '{map_file}'에 저장했습니다.")
                else:
                    logger.warning("지도 데이터를 로드할 수 없습니다.")
                
            except ImportError:
                logger.warning("korea_map 모듈을 로드할 수 없습니다. 네트워크 맵을 시각화하지 않습니다.")
            
        except Exception as e:
            logger.error(f"네트워크 맵 생성 중 오류 발생: {str(e)}")
    
    def create_interactive_dashboard(self, network, result_file):
        """인터랙티브 대시보드 생성 (향후 구현)
        
        Args:
            network (pypsa.Network): 최적화된 PyPSA 네트워크 객체
            result_file (str): 결과 파일 경로
        """
        try:
            # 향후 plotly나 dash로 인터랙티브 대시보드 구현 예정
            pass
        except Exception as e:
            logger.error(f"인터랙티브 대시보드 생성 중 오류 발생: {str(e)}") 