#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
데이터 로더 모듈

엑셀 파일에서 입력 데이터를 로드하고 검증하는 기능을 제공합니다.
"""

import os
import logging
import pandas as pd
import numpy as np
from datetime import datetime

logger = logging.getLogger("PyPSA-HD.DataLoader")

class ExcelDataLoader:
    """엑셀 데이터 로더 클래스"""
    
    def __init__(self, config):
        """초기화 함수
        
        Args:
            config (dict): 설정 정보
        """
        self.config = config
        self.required_sheets = {
            'timeseries': ['start_time', 'end_time', 'frequency'],
            'buses': ['name', 'v_nom', 'carrier', 'x', 'y'],
            'generators': ['name', 'bus', 'carrier', 'p_nom'],
            'lines': ['name', 'bus0', 'bus1', 'carrier', 'x', 'r', 's_nom'],
            'loads': ['name', 'bus', 'p_set'],
            'stores': ['name', 'bus', 'carrier', 'e_nom', 'e_cyclic'],
            'links': ['name', 'bus0', 'bus1', 'efficiency0', 'p_nom'],
            'renewable_patterns': ['hour', 'PV', 'WT']
        }
        
    def load_data(self, input_file):
        """엑셀 파일에서 입력 데이터 로드
        
        Args:
            input_file (str): 입력 데이터 파일 경로
            
        Returns:
            dict: 시트별 데이터 딕셔너리
        """
        logger.info(f"'{input_file}' 파일에서 데이터를 로드합니다.")
        
        try:
            if not os.path.exists(input_file):
                logger.error(f"입력 파일 '{input_file}'을 찾을 수 없습니다.")
                raise FileNotFoundError(f"입력 파일 '{input_file}'을 찾을 수 없습니다.")
            
            # 엑셀 파일 로드
            xls = pd.ExcelFile(input_file)
            input_data = {}
            
            # 각 시트 로드
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(input_file, sheet_name=sheet_name)
                
                # 컬럼명 공백 제거
                df.columns = df.columns.str.strip()
                
                input_data[sheet_name] = df
                logger.debug(f"'{sheet_name}' 시트 로드 완료: {len(df)}행 x {len(df.columns)}열")
            
            # 데이터 검증
            self._validate_data(input_data)
            
            # 데이터 전처리
            input_data = self._preprocess_data(input_data)
            
            logger.info("데이터 로드 및 검증 완료")
            return input_data
            
        except Exception as e:
            logger.error(f"데이터 로드 중 오류 발생: {str(e)}")
            raise
    
    def _validate_data(self, input_data):
        """입력 데이터 유효성 검증
        
        Args:
            input_data (dict): 시트별 데이터 딕셔너리
            
        Raises:
            ValueError: 필수 시트나 컬럼이 없거나 형식이 잘못된 경우
        """
        # 필수 시트 확인
        for sheet_name, required_columns in self.required_sheets.items():
            if sheet_name not in input_data:
                logger.error(f"필수 시트 '{sheet_name}'이 없습니다.")
                raise ValueError(f"필수 시트 '{sheet_name}'이 없습니다.")
            
            df = input_data[sheet_name]
            if df.empty:
                logger.warning(f"'{sheet_name}' 시트가 비어있습니다.")
            
            # 필수 컬럼 확인
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                logger.error(f"'{sheet_name}' 시트에 필수 컬럼이 없습니다: {missing_columns}")
                raise ValueError(f"'{sheet_name}' 시트에 필수 컬럼이 없습니다: {missing_columns}")
        
        # 시간 설정 확인
        if 'timeseries' in input_data:
            timeseries = input_data['timeseries'].iloc[0]
            try:
                pd.to_datetime(timeseries['start_time'])
                pd.to_datetime(timeseries['end_time'])
            except:
                logger.error("잘못된 시간 형식입니다.")
                raise ValueError("잘못된 시간 형식입니다.")
            
            if 'h' not in str(timeseries['frequency']).lower():
                logger.warning("frequency는 'h' 형식을 권장합니다.")
        
        # 버스 참조 확인
        if 'buses' in input_data and 'generators' in input_data:
            available_buses = set(input_data['buses']['name'].astype(str))
            
            # 발전기의 버스 참조 확인
            for i, gen in input_data['generators'].iterrows():
                bus_name = str(gen['bus'])
                if bus_name not in available_buses:
                    logger.warning(f"발전기 '{gen['name']}'의 버스 '{bus_name}'가 정의되지 않았습니다.")
            
            # 부하의 버스 참조 확인
            if 'loads' in input_data:
                for i, load in input_data['loads'].iterrows():
                    bus_name = str(load['bus'])
                    if bus_name not in available_buses:
                        logger.warning(f"부하 '{load['name']}'의 버스 '{bus_name}'가 정의되지 않았습니다.")
    
    def _preprocess_data(self, input_data):
        """입력 데이터 전처리
        
        Args:
            input_data (dict): 시트별 데이터 딕셔너리
            
        Returns:
            dict: 전처리된 데이터 딕셔너리
        """
        # 문자열 필드 공백 제거
        for sheet_name, df in input_data.items():
            string_columns = df.select_dtypes(include=['object']).columns
            for col in string_columns:
                df[col] = df[col].astype(str).str.strip()
            
            input_data[sheet_name] = df
        
        # 재생에너지 패턴 정규화
        if 'renewable_patterns' in input_data:
            patterns_df = input_data['renewable_patterns']
            
            for pattern_type in ['PV', 'WT']:
                if pattern_type in patterns_df.columns:
                    pattern_values = patterns_df[pattern_type].values
                    if np.max(pattern_values) > 0:
                        patterns_df[pattern_type] = pattern_values / np.max(pattern_values)
            
            input_data['renewable_patterns'] = patterns_df
        
        # 타임시리즈 데이터와 부하 패턴 일치 확인 및 조정
        if 'timeseries' in input_data and 'load_patterns' in input_data:
            ts = input_data['timeseries'].iloc[0]
            snapshots = pd.date_range(
                start=ts['start_time'],
                end=ts['end_time'],
                freq=ts['frequency'],
                inclusive='left'
            )
            snapshots_length = len(snapshots)
            
            # 부하 패턴 길이 조정
            load_patterns = input_data['load_patterns'].copy()
            if len(load_patterns) != snapshots_length:
                logger.warning(f"부하 패턴 길이({len(load_patterns)})가 시간 길이({snapshots_length})와 일치하지 않습니다. 조정합니다.")
                
                # 새로운 DataFrame 생성
                new_load_patterns = pd.DataFrame(index=range(1, snapshots_length + 1))
                new_load_patterns['hour'] = range(1, snapshots_length + 1)
                
                # 각 컬럼별 패턴 조정
                for col in load_patterns.columns:
                    if col != 'hour':  # hour 컬럼 제외
                        pattern_values = load_patterns[col].values
                        if len(pattern_values) < snapshots_length:
                            # 부족한 만큼 반복
                            repetitions = snapshots_length // len(pattern_values) + 1
                            extended_pattern = np.tile(pattern_values, repetitions)
                            new_load_patterns[col] = extended_pattern[:snapshots_length]
                        else:
                            # 초과분 제거
                            new_load_patterns[col] = pattern_values[:snapshots_length]
                
                input_data['load_patterns'] = new_load_patterns
        
        return input_data
    
    def adjust_pattern_length(self, pattern_values, required_length):
        """패턴 길이를 필요한 길이에 맞게 조정
        
        Args:
            pattern_values (array): 조정할 패턴 배열
            required_length (int): 필요한 길이
            
        Returns:
            array: 조정된 패턴 배열
        """
        if len(pattern_values) == required_length:
            return pattern_values
        
        # 패턴이 더 짧은 경우
        elif len(pattern_values) < required_length:
            # 부족한 만큼 처음부터 반복
            repetitions = required_length // len(pattern_values) + 1
            extended_pattern = np.tile(pattern_values, repetitions)
            return extended_pattern[:required_length]
        
        # 패턴이 더 긴 경우
        else:
            return pattern_values[:required_length]
    
    def normalize_pattern(self, pattern):
        """발전 패턴을 0~1 사이로 정규화
        
        Args:
            pattern (array): 정규화할 패턴 배열
            
        Returns:
            array: 정규화된 패턴 배열
        """
        if np.max(pattern) > 0:
            return pattern / np.max(pattern)
        return pattern 