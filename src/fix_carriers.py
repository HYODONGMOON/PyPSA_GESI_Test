#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PyPSA-HD 모델의 누락된 carrier 속성을 수정하는 스크립트
"""

import pandas as pd
import os
import sys
import logging
import argparse
from datetime import datetime
import shutil

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("PyPSA-HD-Carrier-Fixer")

def fix_carriers_in_excel(input_file):
    """엑셀 파일에서 모든 구성요소의 carrier 속성을 수정합니다."""
    logger.info(f"'{input_file}' 파일의 carrier 속성 수정을 시작합니다.")
    
    # 파일 존재 여부 확인
    if not os.path.exists(input_file):
        logger.error(f"'{input_file}' 파일을 찾을 수 없습니다.")
        return False
    
    # 백업 파일 생성
    backup_file = f"{input_file.split('.')[0]}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    shutil.copy2(input_file, backup_file)
    logger.info(f"백업 파일이 생성되었습니다: {backup_file}")
    
    try:
        # 엑셀 파일에서 모든 시트 읽기
        with pd.ExcelFile(input_file) as xls:
            sheet_names = xls.sheet_names
            all_data = {}
            
            for sheet in sheet_names:
                all_data[sheet] = pd.read_excel(xls, sheet_name=sheet)
                logger.info(f"'{sheet}' 시트를 로드했습니다. 행 수: {len(all_data[sheet])}")
            
            # carrier 속성 수정
            modified = False
            carrier_components = ['buses', 'lines', 'links', 'generators', 'stores', 'storage_units']
            
            for component in carrier_components:
                if component in all_data:
                    df = all_data[component]
                    if 'carrier' not in df.columns:
                        df['carrier'] = 'default_carrier'
                        all_data[component] = df
                        logger.info(f"'{component}' 시트에 기본 carrier 열을 추가했습니다.")
                        modified = True
                    else:
                        null_carriers = df[df.carrier.isnull()]
                        if not null_carriers.empty:
                            df.loc[df.carrier.isnull(), 'carrier'] = 'default_carrier'
                            all_data[component] = df
                            logger.info(f"'{component}' 시트에서 {len(null_carriers)}개의 누락된 carrier 값을 수정했습니다.")
                            modified = True
            
            # 수정된 경우에만 파일 저장
            if modified:
                with pd.ExcelWriter(input_file) as writer:
                    for sheet_name, df in all_data.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                logger.info(f"수정된 데이터가 '{input_file}'에 저장되었습니다.")
                return True
            else:
                logger.info("모든 carrier 속성이 이미 설정되어 있습니다. 수정이 필요하지 않습니다.")
                return True
                
    except Exception as e:
        logger.error(f"처리 중 오류가 발생했습니다: {str(e)}")
        # 백업에서 복원
        shutil.copy2(backup_file, input_file)
        logger.info(f"오류로 인해 백업에서 파일을 복원했습니다.")
        return False

def main():
    """메인 함수"""
    parser = argparse.ArgumentParser(description="PyPSA-HD 모델의 carrier 속성을 수정합니다.")
    parser.add_argument("--input", dest="input_file", default="simplified_input_data.xlsx",
                      help="입력 엑셀 파일 (기본값: simplified_input_data.xlsx)")
    
    args = parser.parse_args()
    
    logger.info("PyPSA-HD 모델 carrier 속성 수정 시작")
    logger.info(f"Python 버전: {sys.version}")
    logger.info(f"작업 디렉토리: {os.getcwd()}")
    
    # carrier 속성 수정 실행
    success = fix_carriers_in_excel(args.input_file)
    
    if success:
        logger.info("carrier 속성 수정이 성공적으로 완료되었습니다.")
    else:
        logger.error("carrier 속성 수정 중 오류가 발생했습니다.")
        sys.exit(1)

if __name__ == "__main__":
    main() 