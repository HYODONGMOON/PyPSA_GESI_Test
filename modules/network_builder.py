#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
네트워크 빌더 모듈

입력 데이터를 바탕으로 PyPSA 네트워크 객체를 구축합니다.
"""

import logging
import pandas as pd
import numpy as np
import pypsa

logger = logging.getLogger("PyPSA-HD.NetworkBuilder")

class NetworkBuilder:
    """PyPSA 네트워크 빌더 클래스"""
    
    def __init__(self, config):
        """초기화 함수
        
        Args:
            config (dict): 설정 정보
        """
        self.config = config
        self.carriers = {
            'AC': {'name': 'AC', 'co2_emissions': 0},
            'DC': {'name': 'DC', 'co2_emissions': 0},
            'electricity': {'name': 'electricity', 'co2_emissions': 0},
            'coal': {'name': 'coal', 'co2_emissions': 0.9},
            'gas': {'name': 'gas', 'co2_emissions': 0.4},
            'nuclear': {'name': 'nuclear', 'co2_emissions': 0},
            'solar': {'name': 'solar', 'co2_emissions': 0},
            'wind': {'name': 'wind', 'co2_emissions': 0},
            'hydrogen': {'name': 'hydrogen', 'co2_emissions': 0},
            'heat': {'name': 'heat', 'co2_emissions': 0}
        }
    
    def build_network(self, input_data):
        """입력 데이터로부터 PyPSA 네트워크 생성
        
        Args:
            input_data (dict): 시트별 데이터 딕셔너리
            
        Returns:
            pypsa.Network: 구축된 PyPSA 네트워크 객체
        """
        try:
            logger.info("네트워크 구축 시작")
            
            # 네트워크 객체 생성
            network = pypsa.Network()
            
            # carriers 추가
            self._add_carriers(network)
            
            # 시간 설정
            snapshots = self._set_snapshots(network, input_data)
            snapshots_length = len(snapshots) if snapshots is not None else 0
            
            # 버스 추가
            available_buses = self._add_buses(network, input_data)
            
            # 재생에너지 패턴 준비
            renewable_patterns = self._prepare_renewable_patterns(input_data, snapshots_length)
            
            # 발전기 추가
            self._add_generators(network, input_data, available_buses, renewable_patterns)
            
            # 선로 추가
            self._add_lines(network, input_data, available_buses)
            
            # 부하 추가
            self._add_loads(network, input_data, available_buses)
            
            # 저장장치 추가
            self._add_stores(network, input_data, available_buses)
            
            # 링크 추가
            self._add_links(network, input_data, available_buses)
            
            # 글로벌 제약조건 추가
            self._add_global_constraints(network, input_data)
            
            logger.info("네트워크 구축 완료")
            return network
            
        except Exception as e:
            logger.error(f"네트워크 구축 중 오류 발생: {str(e)}")
            raise
    
    def _add_carriers(self, network):
        """에너지 캐리어 추가
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
        """
        for carrier, specs in self.carriers.items():
            network.add("Carrier",
                       name=specs['name'],
                       co2_emissions=specs['co2_emissions'])
        
        logger.debug(f"{len(self.carriers)}개의 캐리어 추가 완료")
    
    def _set_snapshots(self, network, input_data):
        """시간 설정
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            input_data (dict): 입력 데이터
            
        Returns:
            pd.DatetimeIndex: 스냅샷 인덱스 
        """
        if 'timeseries' in input_data and not input_data['timeseries'].empty:
            ts = input_data['timeseries'].iloc[0]
            snapshots = pd.date_range(
                start=ts['start_time'],
                end=ts['end_time'],
                freq=ts['frequency'],
                inclusive='left'
            )
            network.set_snapshots(snapshots)
            logger.debug(f"시간 설정 완료: {snapshots[0]} ~ {snapshots[-1]} ({len(snapshots)}개)")
            return snapshots
        else:
            logger.warning("timeseries 데이터가 없습니다. 기본 시간 설정을 사용합니다.")
            # 기본 시간 설정: 1일, 1시간 간격
            snapshots = pd.date_range(
                start="2024-01-01",
                periods=24,
                freq='h'
            )
            network.set_snapshots(snapshots)
            return snapshots
    
    def _add_buses(self, network, input_data):
        """버스 추가
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            input_data (dict): 입력 데이터
            
        Returns:
            set: 사용 가능한 버스 이름 집합
        """
        available_buses = set()
        
        if 'buses' in input_data and not input_data['buses'].empty:
            for _, bus in input_data['buses'].iterrows():
                bus_name = str(bus['name'])
                available_buses.add(bus_name)
                
                # 기본 파라미터
                params = {
                    'name': bus_name,
                    'v_nom': float(bus['v_nom']),
                    'carrier': str(bus['carrier'])
                }
                
                # 위치 정보 (있는 경우)
                if 'x' in bus and pd.notna(bus['x']):
                    params['x'] = float(bus['x'])
                if 'y' in bus and pd.notna(bus['y']):
                    params['y'] = float(bus['y'])
                
                network.add("Bus", **params)
            
            logger.debug(f"{len(available_buses)}개의 버스 추가 완료")
        else:
            logger.warning("buses 데이터가 없습니다.")
        
        return available_buses
    
    def _prepare_renewable_patterns(self, input_data, snapshots_length):
        """재생에너지 패턴 준비
        
        Args:
            input_data (dict): 입력 데이터
            snapshots_length (int): 시간 길이
            
        Returns:
            dict: 재생에너지 패턴 딕셔너리
        """
        renewable_patterns = {}
        
        if 'renewable_patterns' in input_data and not input_data['renewable_patterns'].empty:
            patterns_df = input_data['renewable_patterns']
            
            # PV 패턴
            if 'PV' in patterns_df.columns:
                pv_pattern = patterns_df['PV'].values
                if snapshots_length > 0:
                    # 길이 조정
                    if len(pv_pattern) != snapshots_length:
                        logger.debug(f"PV 패턴 길이 조정: {len(pv_pattern)} -> {snapshots_length}")
                        repetitions = snapshots_length // len(pv_pattern) + 1
                        pv_pattern = np.tile(pv_pattern, repetitions)[:snapshots_length]
                renewable_patterns['PV_pattern'] = pv_pattern
            
            # WT(풍력) 패턴
            if 'WT' in patterns_df.columns:
                wt_pattern = patterns_df['WT'].values
                if snapshots_length > 0:
                    # 길이 조정
                    if len(wt_pattern) != snapshots_length:
                        logger.debug(f"WT 패턴 길이 조정: {len(wt_pattern)} -> {snapshots_length}")
                        repetitions = snapshots_length // len(wt_pattern) + 1
                        wt_pattern = np.tile(wt_pattern, repetitions)[:snapshots_length]
                renewable_patterns['WT_pattern'] = wt_pattern
            
            logger.debug(f"재생에너지 패턴 준비 완료: {list(renewable_patterns.keys())}")
        else:
            logger.warning("renewable_patterns 데이터가 없습니다.")
        
        return renewable_patterns
    
    def _add_generators(self, network, input_data, available_buses, renewable_patterns):
        """발전기 추가
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            input_data (dict): 입력 데이터
            available_buses (set): 사용 가능한 버스 집합
            renewable_patterns (dict): 재생에너지 패턴
        """
        if 'generators' not in input_data or input_data['generators'].empty:
            logger.warning("generators 데이터가 없습니다.")
            return
        
        generators_added = 0
        
        for _, gen in input_data['generators'].iterrows():
            bus_name = str(gen['bus'])
            
            # 버스 검증
            if bus_name not in available_buses:
                logger.warning(f"발전기 '{gen['name']}'의 버스 '{bus_name}'가 정의되지 않았습니다. 해당 발전기를 건너뜁니다.")
                continue
            
            # 기본 파라미터
            params = {
                'name': str(gen['name']),
                'bus': bus_name,
                'p_nom': float(gen['p_nom']),
                'carrier': str(gen['carrier']) if 'carrier' in gen else self._get_carrier_from_name(gen['name'])
            }
            
            # 선택적 파라미터
            optional_params = [
                'p_nom_extendable', 'p_nom_min', 'p_nom_max', 
                'marginal_cost', 'capital_cost', 'efficiency', 
                'committable', 'ramp_limit_up', 'min_up_time', 
                'start_up_cost', 'lifetime'
            ]
            
            for param in optional_params:
                if param in gen and pd.notna(gen[param]):
                    if param == 'committable':
                        params[param] = bool(gen[param])
                    else:
                        params[param] = float(gen[param])
            
            # 시간별 발전 패턴 설정
            if 'p_max_pu' in gen and pd.notna(gen['p_max_pu']):
                pattern_name = str(gen['p_max_pu'])
                if pattern_name in renewable_patterns:
                    p_max_pu = renewable_patterns[pattern_name]
                    network.add("Generator", **params)
                    network.generators_t.p_max_pu[params['name']] = p_max_pu
                    generators_added += 1
                else:
                    logger.warning(f"발전기 '{gen['name']}'의 패턴 '{pattern_name}'을 찾을 수 없습니다.")
                    network.add("Generator", **params)
                    generators_added += 1
            else:
                network.add("Generator", **params)
                generators_added += 1
        
        logger.debug(f"{generators_added}개의 발전기 추가 완료")
    
    def _get_carrier_from_name(self, gen_name):
        """발전기 이름으로부터 캐리어 유추
        
        Args:
            gen_name (str): 발전기 이름
            
        Returns:
            str: 캐리어 이름
        """
        name = str(gen_name).lower()
        if 'pv' in name:
            return 'solar'
        elif 'wt' in name or 'wind' in name:
            return 'wind'
        elif 'nuclear' in name:
            return 'nuclear'
        elif 'coal' in name:
            return 'coal'
        elif 'gas' in name:
            return 'gas'
        else:
            return 'electricity'
    
    def _add_lines(self, network, input_data, available_buses):
        """선로 추가
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            input_data (dict): 입력 데이터
            available_buses (set): 사용 가능한 버스 집합
        """
        if 'lines' not in input_data or input_data['lines'].empty:
            logger.warning("lines 데이터가 없습니다.")
            return
        
        lines_added = 0
        
        for _, line in input_data['lines'].iterrows():
            bus0 = str(line['bus0'])
            bus1 = str(line['bus1'])
            
            # 버스 검증
            if bus0 not in available_buses:
                logger.warning(f"선로 '{line['name']}'의 bus0 '{bus0}'가 정의되지 않았습니다. 해당 선로를 건너뜁니다.")
                continue
            if bus1 not in available_buses:
                logger.warning(f"선로 '{line['name']}'의 bus1 '{bus1}'가 정의되지 않았습니다. 해당 선로를 건너뜁니다.")
                continue
            
            # 기본 파라미터
            params = {
                'name': str(line['name']),
                'bus0': bus0,
                'bus1': bus1,
                'carrier': str(line['carrier']) if 'carrier' in line else 'AC'
            }
            
            # 선택적 파라미터
            optional_params = ['x', 'r', 's_nom', 's_nom_extendable', 's_nom_min', 's_nom_max', 'length', 'capital_cost']
            
            for param in optional_params:
                if param in line and pd.notna(line[param]):
                    params[param] = float(line[param])
            
            network.add("Line", **params)
            lines_added += 1
        
        logger.debug(f"{lines_added}개의 선로 추가 완료")
    
    def _add_loads(self, network, input_data, available_buses):
        """부하 추가
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            input_data (dict): 입력 데이터
            available_buses (set): 사용 가능한 버스 집합
        """
        if 'loads' not in input_data or input_data['loads'].empty:
            logger.warning("loads 데이터가 없습니다.")
            return
        
        loads_added = 0
        snapshots = network.snapshots
        snapshots_length = len(snapshots)
        
        for _, load in input_data['loads'].iterrows():
            bus_name = str(load['bus'])
            
            # 버스 검증
            if bus_name not in available_buses:
                logger.warning(f"부하 '{load['name']}'의 버스 '{bus_name}'가 정의되지 않았습니다. 해당 부하를 건너뜁니다.")
                continue
            
            # 기본 파라미터
            params = {
                'name': str(load['name']),
                'bus': bus_name,
                'p_set': float(load['p_set'])
            }
            
            network.add("Load", **params)
            loads_added += 1
            
            # 시간별 패턴 적용 (load_patterns 시트가 있는 경우)
            if 'load_patterns' in input_data and not input_data['load_patterns'].empty:
                load_patterns = input_data['load_patterns']
                load_name = str(load['name'])
                
                if load_name in load_patterns.columns:
                    pattern = load_patterns[load_name].values
                    
                    # 패턴 길이 조정
                    if len(pattern) != snapshots_length:
                        repetitions = snapshots_length // len(pattern) + 1
                        pattern = np.tile(pattern, repetitions)[:snapshots_length]
                    
                    # 패턴 적용: p_set * pattern
                    network.loads_t.p_set[load_name] = float(load['p_set']) * pattern
                    logger.debug(f"부하 '{load_name}'에 시간별 패턴 적용 완료")
        
        logger.debug(f"{loads_added}개의 부하 추가 완료")
    
    def _add_stores(self, network, input_data, available_buses):
        """저장장치 추가
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            input_data (dict): 입력 데이터
            available_buses (set): 사용 가능한 버스 집합
        """
        if 'stores' not in input_data or input_data['stores'].empty:
            logger.warning("stores 데이터가 없습니다.")
            return
        
        stores_added = 0
        
        for _, store in input_data['stores'].iterrows():
            bus_name = str(store['bus'])
            
            # 버스 검증
            if bus_name not in available_buses:
                logger.warning(f"저장장치 '{store['name']}'의 버스 '{bus_name}'가 정의되지 않았습니다. 해당 저장장치를 건너뜁니다.")
                continue
            
            # 기본 파라미터
            params = {
                'name': str(store['name']),
                'bus': bus_name,
                'carrier': str(store['carrier']) if 'carrier' in store else 'electricity',
                'e_nom': float(store['e_nom']),
                'e_cyclic': bool(store['e_cyclic']) if 'e_cyclic' in store else False
            }
            
            # 선택적 파라미터
            optional_params = [
                'e_nom_extendable', 'e_nom_min', 'e_nom_max', 
                'standing_loss', 'efficiency_store', 'efficiency_dispatch', 
                'e_initial', 'capital_cost'
            ]
            
            for param in optional_params:
                if param in store and pd.notna(store[param]):
                    params[param] = float(store[param]) if param != 'e_nom_extendable' else bool(store[param])
            
            network.add("Store", **params)
            stores_added += 1
        
        logger.debug(f"{stores_added}개의 저장장치 추가 완료")
    
    def _add_links(self, network, input_data, available_buses):
        """링크 추가
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            input_data (dict): 입력 데이터
            available_buses (set): 사용 가능한 버스 집합
        """
        if 'links' not in input_data or input_data['links'].empty:
            logger.warning("links 데이터가 없습니다.")
            return
        
        links_added = 0
        
        for _, link in input_data['links'].iterrows():
            # 버스 검증
            required_buses = {}
            for bus_col in ['bus0', 'bus1', 'bus2', 'bus3']:
                if bus_col in link and pd.notna(link[bus_col]):
                    bus_name = str(link[bus_col])
                    if bus_name not in available_buses:
                        logger.warning(f"링크 '{link['name']}'의 {bus_col} '{bus_name}'가 정의되지 않았습니다.")
                        break
                    required_buses[bus_col] = bus_name
            
            # 필수 버스가 없으면 건너뛰기
            if 'bus0' not in required_buses:
                logger.warning(f"링크 '{link['name']}'에 필수 버스 'bus0'가 없습니다. 해당 링크를 건너뜁니다.")
                continue
            
            # 기본 파라미터
            params = {
                'name': str(link['name']),
                'bus0': required_buses['bus0']
            }
            
            # p_nom 설정 (필수)
            if 'p_nom' in link and pd.notna(link['p_nom']):
                params['p_nom'] = float(link['p_nom'])
            else:
                logger.warning(f"링크 '{link['name']}'에 필수 파라미터 'p_nom'이 없습니다. 기본값 100을 사용합니다.")
                params['p_nom'] = 100.0
            
            # 추가 버스 연결
            for bus_col in ['bus1', 'bus2', 'bus3']:
                if bus_col in required_buses:
                    params[bus_col] = required_buses[bus_col]
            
            # 효율 설정
            if 'efficiency0' in link and pd.notna(link['efficiency0']):
                params['efficiency'] = float(link['efficiency0'])
            if 'bus3' in required_buses and 'efficiency1' in link and pd.notna(link['efficiency1']):
                params['efficiency2'] = float(link['efficiency1'])
            
            # 운영 제약 파라미터
            optional_params = [
                'p_nom_extendable', 'p_nom_min', 'p_nom_max', 
                'capital_cost', 'marginal_cost', 'committable', 
                'p_min_pu', 'p_max_pu', 'ramp_limit_up', 'ramp_limit_down',
                'min_up_time', 'min_down_time', 'start_up_cost', 'shut_down_cost'
            ]
            
            for param in optional_params:
                if param in link and pd.notna(link[param]):
                    if param == 'committable':
                        params[param] = bool(link[param])
                    elif param in ['p_min_pu', 'p_max_pu']:
                        params[param] = float(link[param])
                    else:
                        params[param] = float(link[param])
            
            network.add("Link", **params)
            links_added += 1
        
        logger.debug(f"{links_added}개의 링크 추가 완료")
    
    def _add_global_constraints(self, network, input_data):
        """글로벌 제약조건 추가
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            input_data (dict): 입력 데이터
        """
        if 'constraints' not in input_data or input_data['constraints'].empty:
            logger.debug("constraints 데이터가 없습니다.")
            return
        
        constraints_added = 0
        
        for _, constraint in input_data['constraints'].iterrows():
            name = str(constraint['name'])
            
            # 기본 파라미터
            params = {
                'name': name,
                'sense': str(constraint['sense']),
                'constant': float(constraint['constant']) if 'constant' in constraint else 0.0
            }
            
            # 제약조건 타입 확인
            constraint_type = str(constraint['type']) if 'type' in constraint else None
            
            if constraint_type == 'primary_energy':
                # 주 에너지 제약조건 (예: CO2 제한)
                if 'carrier_attribute' in constraint and pd.notna(constraint['carrier_attribute']):
                    attribute = str(constraint['carrier_attribute'])
                    network.add("GlobalConstraint", type='primary_energy', carrier_attribute=attribute, **params)
                    constraints_added += 1
                else:
                    logger.warning(f"제약조건 '{name}'에 필수 파라미터 'carrier_attribute'가 없습니다.")
            
            elif constraint_type == 'transmission_volume_expansion_limit':
                # 송전 확장 제약조건
                network.add("GlobalConstraint", type='transmission_volume_expansion_limit', **params)
                constraints_added += 1
            
            else:
                logger.warning(f"알 수 없는 제약조건 타입: {constraint_type}")
        
        if constraints_added > 0:
            logger.debug(f"{constraints_added}개의 제약조건 추가 완료") 