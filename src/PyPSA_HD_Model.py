#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PyPSA-HD 전력 시스템 최적화 모델

엑셀 기반 입력을 강화한 PyPSA 기반 전력 시스템 최적화 모델입니다.
"""

import os
import sys
import time
import logging
import pandas as pd
import numpy as np
import traceback
import multiprocessing
from datetime import datetime

# PyPSA 라이브러리 임포트
import pypsa

# 내부 모듈 임포트
from modules.data_loader import ExcelDataLoader
from modules.network_builder import NetworkBuilder
from modules.optimizer import PypsaOptimizer
from modules.result_processor import ResultProcessor
from modules.visualization import Visualizer

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("pypsa_hd.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("PyPSA-HD")

class PypsaHDModel:
    """PyPSA-HD 전력 시스템 최적화 모델 클래스"""
    
    def __init__(self, config_file=None):
        """초기화 함수
        
        Args:
            config_file (str, optional): 설정 파일 경로. 지정하지 않으면 기본 설정 사용.
        """
        self.config = self._load_config(config_file)
        self.data_loader = ExcelDataLoader(self.config)
        self.network_builder = NetworkBuilder(self.config)
        self.optimizer = PypsaOptimizer(self.config)
        self.result_processor = ResultProcessor(self.config)
        self.visualizer = Visualizer(self.config)
        self.network = None
        
        logger.info("PyPSA-HD 모델이 초기화되었습니다.")
    
    def _load_config(self, config_file):
        """설정 파일 로드
        
        Args:
            config_file (str): 설정 파일 경로
            
        Returns:
            dict: 설정 정보
        """
        # 기본 설정값
        default_config = {
            'input_file': "input_data.xlsx",
            'output_dir': "results",
            'solver': {
                'name': 'cplex',
                'threads': multiprocessing.cpu_count(),
                'options': {
                    'lpmethod': 4,  # Barrier method
                    'barrier.algorithm': 3,
                    'parallel': 1,
                    'solutiontype': 2
                }
            },
            'visualization': {
                'enabled': True,
                'map_file': 'korea_map.png'
            }
        }
        
        # 설정 파일이 제공된 경우 해당 파일에서 설정 로드
        if config_file and os.path.exists(config_file):
            try:
                user_config = pd.read_excel(config_file, sheet_name='config').set_index('parameter').to_dict()['value']
                # 기본 설정에 사용자 설정 적용
                for key, value in user_config.items():
                    if isinstance(value, str) and value.lower() in ['true', 'false']:
                        value = value.lower() == 'true'
                    default_config[key] = value
                logger.info(f"설정 파일 '{config_file}'에서 설정을 로드했습니다.")
            except Exception as e:
                logger.warning(f"설정 파일 로드 중 오류 발생: {str(e)}. 기본 설정을 사용합니다.")
        else:
            logger.info("설정 파일이 없습니다. 기본 설정을 사용합니다.")
            
        # 결과 디렉토리 생성
        os.makedirs(default_config['output_dir'], exist_ok=True)
            
        return default_config
        
    def run(self):
        """모델 실행"""
        try:
            start_time = time.time()
            logger.info("모델 실행을 시작합니다.")
            
            # 1. 데이터 로드
            logger.info("입력 데이터를 로드합니다...")
            input_data = self.data_loader.load_data(self.config['input_file'])
            
            # 2. 네트워크 생성
            logger.info("네트워크 모델을 생성합니다...")
            self.network = self.network_builder.build_network(input_data)
            
            # 3. 모델 최적화
            logger.info("네트워크 최적화를 시작합니다...")
            optimization_success = self.optimizer.optimize(self.network)
            
            if not optimization_success:
                logger.error("최적화에 실패했습니다.")
                return False
            
            # 4. 결과 처리
            logger.info("최적화 결과를 처리합니다...")
            result_file = self.result_processor.process_results(self.network)
            
            # 5. 시각화 (설정에서 활성화된 경우)
            if self.config['visualization']['enabled']:
                logger.info("결과를 시각화합니다...")
                self.visualizer.visualize_results(self.network, input_data)
            
            elapsed_time = time.time() - start_time
            logger.info(f"모델 실행이 완료되었습니다. 소요 시간: {elapsed_time:.2f}초")
            logger.info(f"결과 파일: {result_file}")
            
            return True
            
        except Exception as e:
            logger.error(f"모델 실행 중 오류 발생: {str(e)}")
            traceback.print_exc()
            return False
    
    def get_network(self):
        """최적화된 네트워크 객체 반환"""
        return self.network

def main():
    """명령행에서 실행할 경우의 메인 함수"""
    import argparse
    
    # 명령행 인수 파싱
    parser = argparse.ArgumentParser(description='PyPSA-HD 전력 시스템 최적화 모델')
    parser.add_argument('--config', type=str, help='설정 파일 경로')
    parser.add_argument('--input', type=str, help='입력 데이터 파일 경로')
    args = parser.parse_args()
    
    # 모델 초기화
    model = PypsaHDModel(config_file=args.config)
    
    # 입력 파일 경로가 명령행에서 제공된 경우 설정 업데이트
    if args.input:
        model.config['input_file'] = args.input
    
    # 모델 실행
    success = model.run()
    
    # 종료 코드 설정 (성공: 0, 실패: 1)
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main() 