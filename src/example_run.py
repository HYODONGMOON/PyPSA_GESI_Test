#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PyPSA-HD 모델 예제 실행 스크립트

PyPSA-HD 모델을 실행하는 예제 스크립트입니다.
"""

import os
import sys
import logging
import argparse
from datetime import datetime

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(f"pypsa_hd_example_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("PyPSA-HD-Example")

# PyPSA-HD 모델 클래스 임포트
from PyPSA_HD_Model import PypsaHDModel

def run_example(input_file=None, config_file=None):
    """모델 예제 실행
    
    Args:
        input_file (str, optional): 입력 데이터 파일 경로
        config_file (str, optional): 설정 파일 경로
        
    Returns:
        bool: 실행 성공 여부
    """
    try:
        logger.info("PyPSA-HD 모델 예제 실행을 시작합니다.")
        
        # 기본 입력 파일 설정
        if input_file is None:
            input_file = "input_data.xlsx"
            logger.info(f"입력 파일이 지정되지 않아 기본 파일({input_file})을 사용합니다.")
        
        # 입력 파일 존재 확인
        if not os.path.exists(input_file):
            logger.error(f"입력 파일 '{input_file}'을 찾을 수 없습니다.")
            return False
        
        # 모델 초기화
        model = PypsaHDModel(config_file=config_file)
        
        # 입력 파일 경로 설정
        model.config['input_file'] = input_file
        
        # 모델 실행
        success = model.run()
        
        if success:
            logger.info("모델 실행이 성공적으로 완료되었습니다.")
            return True
        else:
            logger.error("모델 실행이 실패했습니다.")
            return False
        
    except Exception as e:
        logger.error(f"예제 실행 중 오류 발생: {str(e)}")
        return False

def main():
    """메인 함수"""
    parser = argparse.ArgumentParser(description='PyPSA-HD 모델 예제 실행')
    parser.add_argument('--input', type=str, help='입력 데이터 파일 경로')
    parser.add_argument('--config', type=str, help='설정 파일 경로')
    
    args = parser.parse_args()
    
    success = run_example(input_file=args.input, config_file=args.config)
    
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main() 