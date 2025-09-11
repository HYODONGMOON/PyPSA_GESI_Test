#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
결과 처리 모듈

최적화 결과를 처리하고 엑셀 파일로 저장합니다.
"""

import os
import logging
import pandas as pd
import numpy as np
from datetime import datetime

logger = logging.getLogger("PyPSA-HD.ResultProcessor")

class ResultProcessor:
    """결과 처리 클래스"""
    
    def __init__(self, config):
        """초기화 함수
        
        Args:
            config (dict): 설정 정보
        """
        self.config = config
        self.output_dir = config.get('output_dir', 'results')
        os.makedirs(self.output_dir, exist_ok=True)
    
    def process_results(self, network):
        """최적화 결과를 처리하고 저장
        
        Args:
            network (pypsa.Network): 최적화된 PyPSA 네트워크 객체
            
        Returns:
            str: 결과 파일 경로
        """
        try:
            if not hasattr(network, 'objective') or network.objective is None:
                logger.warning("최적화 결과가 없어 저장할 수 없습니다.")
                return None
            
            # 결과 데이터 추출
            results = self._extract_results(network)
            
            # 결과 파일 경로 생성
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f'results_{current_time}.xlsx'
            result_file = os.path.join(self.output_dir, filename)
            
            # 결과 저장
            self._save_results_to_excel(results, result_file)
            
            logger.info(f"결과를 '{result_file}'에 저장했습니다.")
            return result_file
            
        except Exception as e:
            logger.error(f"결과 처리 중 오류 발생: {str(e)}")
            return None
    
    def _extract_results(self, network):
        """네트워크에서 주요 결과 추출
        
        Args:
            network (pypsa.Network): 최적화된 PyPSA 네트워크 객체
            
        Returns:
            dict: 결과 데이터 딕셔너리
        """
        logger.info("최적화 결과 추출 중...")
        
        results = {
            'summary': self._create_summary(network),
            'generator_output': network.generators_t.p,
            'generator_info': self._get_generator_info(network),
            'bus_info': self._get_bus_info(network),
            'bus_prices': network.buses_t.marginal_price,
            'line_info': self._get_line_info(network),
            'line_flow': network.lines_t.p0 if not network.lines.empty else None,
            'link_info': self._get_link_info(network),
            'link_flow': network.links_t.p0 if not network.links.empty else None,
            'load_info': self._get_load_info(network),
            'load_values': network.loads_t.p_set,
        }
        
        # 저장장치 결과 (있는 경우)
        if not network.stores.empty:
            results['store_info'] = self._get_store_info(network)
            results['store_energy'] = network.stores_t.e
            results['store_power'] = network.stores_t.p
        
        return results
    
    def _create_summary(self, network):
        """최적화 요약 정보 생성
        
        Args:
            network (pypsa.Network): 최적화된 PyPSA 네트워크 객체
            
        Returns:
            pd.DataFrame: 요약 정보 데이터프레임
        """
        # 기본 정보
        summary_data = {
            'Parameter': [
                '총 비용', 
                '상태', 
                '발전기 수', 
                '버스 수', 
                '선로 수',
                '링크 수',
                '저장장치 수',
                '부하 수',
                '타임스텝 수',
            ],
            'Value': [
                network.objective,
                'Optimal',
                len(network.generators),
                len(network.buses),
                len(network.lines),
                len(network.links),
                len(network.stores),
                len(network.loads),
                len(network.snapshots)
            ]
        }
        
        summary_df = pd.DataFrame(summary_data)
        
        # 발전량 및 부하량 총합 계산
        if not network.generators_t.p.empty:
            total_generation = network.generators_t.p.sum().sum()
            summary_df = pd.concat([
                summary_df,
                pd.DataFrame({
                    'Parameter': ['총 발전량'],
                    'Value': [total_generation]
                })
            ], ignore_index=True)
        
        if not network.loads_t.p.empty:
            total_load = network.loads_t.p.sum().sum()
            summary_df = pd.concat([
                summary_df,
                pd.DataFrame({
                    'Parameter': ['총 소비량'],
                    'Value': [total_load]
                })
            ], ignore_index=True)
        
        return summary_df
    
    def _get_generator_info(self, network):
        """발전기 정보 추출
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            
        Returns:
            pd.DataFrame: 발전기 정보 데이터프레임
        """
        if network.generators.empty:
            return pd.DataFrame()
            
        # 기본 발전기 정보
        gen_info = network.generators.copy()
        
        # 발전량 추가
        if not network.generators_t.p.empty:
            gen_info['p_mean'] = network.generators_t.p.mean()
            gen_info['p_max'] = network.generators_t.p.max()
            gen_info['p_min'] = network.generators_t.p.min()
            gen_info['p_sum'] = network.generators_t.p.sum()
        
        return gen_info
    
    def _get_bus_info(self, network):
        """버스 정보 추출
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            
        Returns:
            pd.DataFrame: 버스 정보 데이터프레임
        """
        if network.buses.empty:
            return pd.DataFrame()
            
        bus_info = network.buses.copy()
        
        # 한계가격 정보 추가
        if not network.buses_t.marginal_price.empty:
            bus_info['price_mean'] = network.buses_t.marginal_price.mean()
            bus_info['price_max'] = network.buses_t.marginal_price.max()
            bus_info['price_min'] = network.buses_t.marginal_price.min()
        
        return bus_info
    
    def _get_line_info(self, network):
        """선로 정보 추출
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            
        Returns:
            pd.DataFrame: 선로 정보 데이터프레임
        """
        if network.lines.empty:
            return pd.DataFrame()
            
        line_info = network.lines.copy()
        
        # 조류 정보 추가
        if not network.lines_t.p0.empty:
            line_info['p_mean'] = network.lines_t.p0.mean()
            line_info['p_max'] = network.lines_t.p0.max()
            line_info['p_min'] = network.lines_t.p0.min()
            line_info['loading_max'] = (network.lines_t.p0.abs().max() / line_info['s_nom']) * 100
        
        return line_info
    
    def _get_link_info(self, network):
        """링크 정보 추출
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            
        Returns:
            pd.DataFrame: 링크 정보 데이터프레임
        """
        if network.links.empty:
            return pd.DataFrame()
            
        link_info = network.links.copy()
        
        # 조류 정보 추가
        if not network.links_t.p0.empty:
            link_info['p_mean'] = network.links_t.p0.mean()
            link_info['p_max'] = network.links_t.p0.max()
            link_info['p_min'] = network.links_t.p0.min()
            link_info['loading_max'] = (network.links_t.p0.abs().max() / link_info['p_nom']) * 100
        
        return link_info
    
    def _get_store_info(self, network):
        """저장장치 정보 추출
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            
        Returns:
            pd.DataFrame: 저장장치 정보 데이터프레임
        """
        if network.stores.empty:
            return pd.DataFrame()
            
        store_info = network.stores.copy()
        
        # 에너지 정보 추가
        if hasattr(network, 'stores_t') and not network.stores_t.empty:
            if not network.stores_t.e.empty:
                store_info['e_mean'] = network.stores_t.e.mean()
                store_info['e_max'] = network.stores_t.e.max()
                store_info['e_min'] = network.stores_t.e.min()
            
            # 충방전 정보
            if not network.stores_t.p.empty:
                store_info['p_mean'] = network.stores_t.p.mean()
                store_info['p_max'] = network.stores_t.p.max()
                store_info['p_min'] = network.stores_t.p.min()
                
                # 충방전 사이클 계산
                for store in network.stores.index:
                    if store in network.stores_t.p:
                        p_store = network.stores_t.p[store]
                        # 충전->방전 전환 횟수 계산 (충전: 양수, 방전: 음수)
                        sign_changes = ((p_store.shift() * p_store) < 0).sum()
                        store_info.loc[store, 'cycles'] = sign_changes / 2
        
        return store_info
    
    def _get_load_info(self, network):
        """부하 정보 추출
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            
        Returns:
            pd.DataFrame: 부하 정보 데이터프레임
        """
        if network.loads.empty:
            return pd.DataFrame()
            
        load_info = network.loads.copy()
        
        # 부하 패턴 정보 추가
        if not network.loads_t.p_set.empty:
            load_info['p_mean'] = network.loads_t.p_set.mean()
            load_info['p_max'] = network.loads_t.p_set.max()
            load_info['p_min'] = network.loads_t.p_set.min()
            load_info['p_sum'] = network.loads_t.p_set.sum()
        
        return load_info
    
    def _save_results_to_excel(self, results, filename):
        """결과를 엑셀 파일로 저장
        
        Args:
            results (dict): 결과 데이터 딕셔너리
            filename (str): 저장할 파일 경로
            
        Returns:
            bool: 저장 성공 여부
        """
        try:
            logger.info(f"결과를 {filename}에 저장합니다...")
            
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # 요약 정보
                if 'summary' in results and not results['summary'].empty:
                    results['summary'].to_excel(writer, sheet_name='Summary', index=False)
                
                # 발전기 출력 및 정보
                if 'generator_output' in results and not results['generator_output'].empty:
                    results['generator_output'].to_excel(writer, sheet_name='Generator_Output')
                
                if 'generator_info' in results and not results['generator_info'].empty:
                    results['generator_info'].to_excel(writer, sheet_name='Generator_Info')
                
                # 버스 정보 및 가격
                if 'bus_info' in results and not results['bus_info'].empty:
                    results['bus_info'].to_excel(writer, sheet_name='Bus_Info')
                
                if 'bus_prices' in results and not results['bus_prices'].empty:
                    results['bus_prices'].to_excel(writer, sheet_name='Bus_Prices')
                
                # 선로 정보 및 조류
                if 'line_info' in results and not results['line_info'].empty:
                    results['line_info'].to_excel(writer, sheet_name='Line_Info')
                
                if 'line_flow' in results and results['line_flow'] is not None:
                    results['line_flow'].to_excel(writer, sheet_name='Line_Flow')
                
                # 링크 정보 및 조류
                if 'link_info' in results and not results['link_info'].empty:
                    results['link_info'].to_excel(writer, sheet_name='Link_Info')
                
                if 'link_flow' in results and results['link_flow'] is not None:
                    results['link_flow'].to_excel(writer, sheet_name='Link_Flow')
                
                # 부하 정보 및 값
                if 'load_info' in results and not results['load_info'].empty:
                    results['load_info'].to_excel(writer, sheet_name='Load_Info')
                
                if 'load_values' in results and not results['load_values'].empty:
                    results['load_values'].to_excel(writer, sheet_name='Load_Values')
                
                # 저장장치 정보 및 값
                if 'store_info' in results and not results['store_info'].empty:
                    results['store_info'].to_excel(writer, sheet_name='Storage_Info')
                
                if 'store_energy' in results and not results['store_energy'].empty:
                    results['store_energy'].to_excel(writer, sheet_name='Storage_Energy')
                
                if 'store_power' in results and not results['store_power'].empty:
                    results['store_power'].to_excel(writer, sheet_name='Storage_Power')
            
            logger.info(f"결과 저장 완료: {filename}")
            return True
            
        except Exception as e:
            logger.error(f"결과 저장 중 오류 발생: {str(e)}")
            return False 