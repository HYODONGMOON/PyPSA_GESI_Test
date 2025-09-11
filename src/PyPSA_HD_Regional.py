#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
PyPSA-HD 지역 기반 인터페이스

PyPSA-HD 모델을 지역 기반으로 실행하기 위한 통합 인터페이스입니다.
"""

import os
import sys
import argparse
import pandas as pd
import numpy as np
from datetime import datetime

# 지역 모듈 임포트
from regional_selector import RegionalSelector
from regional_data_manager import RegionalDataManager
from create_regional_excel_template import create_excel_template
from process_regional_excel import RegionalExcelProcessor

# PyPSA_GUI 모듈 임포트 (순환 참조 방지를 위해 주석 처리)
# import PyPSA_GUI

class PyPSA_HD_Regional:
    """PyPSA-HD 지역 기반 인터페이스 클래스"""
    
    def __init__(self):
        """초기화 함수"""
        self.region_selector = RegionalSelector()
        self.data_manager = RegionalDataManager(self.region_selector)
        
        # 기본 파일 경로
        self.template_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'interface.xlsx'))
        self.output_path = "integrated_input_data.xlsx"
        self.result_path = None
    
    def create_template(self):
        """Excel 템플릿 생성"""
        print("Excel 템플릿 생성 중...")
        create_excel_template()
        print(f"Excel 템플릿이 '{self.template_path}'에 생성되었습니다.")
        print("이 Excel 파일을 열어 데이터를 입력한 후, 'process_template' 명령을 실행하세요.")
        return True
    
    def process_template(self, template_path=None):
        """템플릿 처리 및 모델 실행"""
        if template_path:
            self.template_path = template_path
        
        print(f"템플릿 '{self.template_path}' 처리 중...")
        processor = RegionalExcelProcessor(self.template_path)
        processor.output_path = self.output_path
        
        # 템플릿 처리 및 모델 실행
        if not processor.process_excel():
            print("템플릿 처리 중 오류가 발생했습니다.")
            return False
        
        self.result_path = processor.output_path
        return True
    
    def save_region_data(self, template_path=None, region_code=None):
        """특정 지역의, 또는 모든 선택된 지역의 데이터 저장
        
        Args:
            template_path (str): 템플릿 파일 경로 (없으면 기본값 사용)
            region_code (str): 저장할 지역 코드 (없으면 모든 선택된 지역 처리)
            
        Returns:
            bool: 성공 여부
        """
        if template_path:
            self.template_path = template_path
            
        try:
            print(f"지역 데이터 저장 중...")
            import openpyxl
            
            # 엑셀 파일 열기
            if not os.path.exists(self.template_path):
                print(f"오류: 템플릿 파일 '{self.template_path}'가 존재하지 않습니다.")
                return False
                
            wb = openpyxl.load_workbook(self.template_path, data_only=True)
            processor = RegionalExcelProcessor(self.template_path)
            
            # 기존 통합 데이터가 있으면 로드 (없으면 새로 생성)
            existing_data = {}
            if os.path.exists(self.output_path):
                try:
                    print(f"기존 통합 데이터 '{self.output_path}'를 로드합니다.")
                    # 각 구성요소(buses, generators 등)별로 DataFrame 로드
                    with pd.ExcelFile(self.output_path) as xls:
                        for sheet_name in xls.sheet_names:
                            existing_data[sheet_name] = pd.read_excel(self.output_path, sheet_name=sheet_name)
                except Exception as e:
                    print(f"기존 데이터 로드 중 오류 발생: {str(e)}")
                    existing_data = {}
            
            # 특정 지역만 저장하는 경우
            if region_code:
                # 해당 지역이 유효한지 확인
                if region_code not in processor.region_selector.regions:
                    print(f"오류: 유효하지 않은 지역 코드 '{region_code}'")
                    return False
                    
                # 해당 지역만 선택하도록 설정
                processor.selected_regions = [region_code]
                print(f"'{region_code}' 지역 데이터만 저장합니다.")
                
                # 지역 데이터 읽기
                processor.data_manager.initialize_region(region_code)
                processor.read_regional_data(wb)
                
                # 통합 데이터 생성
                new_data = processor.data_manager.merge_data()
                if not new_data:
                    print(f"오류: '{region_code}' 지역 데이터를 생성할 수 없습니다.")
                    return False
                
                # 기존 데이터와 병합 (중복 제거)
                merged_data = {}
                for component, df in new_data.items():
                    if not df.empty:
                        if component in existing_data and not existing_data[component].empty:
                            # 지역 필터링 - 동일 지역의 기존 데이터 제거
                            if 'region' in existing_data[component].columns:
                                # region_code와 일치하는 행 제거
                                filtered_df = existing_data[component][existing_data[component]['region'] != region_code]
                                # 새 데이터와 병합
                                merged_data[component] = pd.concat([filtered_df, df], ignore_index=True)
                                print(f"  - {component}: 기존 {len(filtered_df)}개 + 새로운 {len(df)}개 = 총 {len(merged_data[component])}개 항목")
                            else:
                                # region 필드가 없는 경우 그대로 병합
                                merged_data[component] = pd.concat([existing_data[component], df], ignore_index=True)
                                print(f"  - {component}: region 필드 없음, 그대로 병합")
                        else:
                            # 기존 데이터가 없거나 비어있으면 새 데이터 사용
                            merged_data[component] = df
                            print(f"  - {component}: 새로운 데이터 {len(df)}개 항목 추가")
                
                # 시간 설정과 패턴 데이터는 그대로 유지
                if 'timeseries' in existing_data:
                    merged_data['timeseries'] = existing_data['timeseries']
                    print("  - 기존 시간 설정 유지")
                
                if 'renewable_patterns' in existing_data:
                    merged_data['renewable_patterns'] = existing_data['renewable_patterns']
                    print("  - 기존 재생에너지 패턴 유지")
                
                if 'load_patterns' in existing_data:
                    merged_data['load_patterns'] = existing_data['load_patterns']
                    print("  - 기존 부하 패턴 유지")
                
                if 'constraints' in existing_data:
                    merged_data['constraints'] = existing_data['constraints']
                    print("  - 기존 제약사항 유지")
                
                # 병합된 데이터 저장
                with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
                    for component, df in merged_data.items():
                        if not df.empty:
                            df.to_excel(writer, sheet_name=component, index=False)
                
                print(f"지역 '{region_code}' 데이터가 '{self.output_path}'에 통합되었습니다.")
            else:
                # 모든 선택된 지역 읽기
                if not processor.read_selected_regions(wb):
                    print("오류: 선택된 지역을 읽을 수 없습니다.")
                    return False
                    
                if not processor.selected_regions:
                    print("오류: 선택된 지역이 없습니다.")
                    return False
                    
                print(f"선택된 모든 지역 데이터를 저장합니다: {processor.selected_regions}")
                
                # 지역 데이터 읽기
                processor.read_regional_data(wb)
                
                # 시간 설정 및 패턴 데이터 읽기
                processor.read_timeseries(wb)
                processor.read_patterns(wb)
                
                # 통합 데이터 생성
                new_data = processor.data_manager.merge_data()
                if not new_data:
                    print("오류: 통합 데이터를 생성할 수 없습니다.")
                    return False
                
                # 기존 데이터와 병합 (중복 제거)
                merged_data = {}
                
                # 기존 데이터에서 선택된 지역을 제외한 데이터만 유지
                for component, df in new_data.items():
                    if not df.empty:
                        if component in existing_data and not existing_data[component].empty:
                            # 지역 필터링 - 선택된 지역의 기존 데이터 제거
                            if 'region' in existing_data[component].columns:
                                # 선택된 지역과 일치하지 않는 행만 필터링
                                filtered_df = existing_data[component][~existing_data[component]['region'].isin(processor.selected_regions)]
                                # 새 데이터와 병합
                                merged_data[component] = pd.concat([filtered_df, df], ignore_index=True)
                                print(f"  - {component}: 기존 {len(filtered_df)}개 + 새로운 {len(df)}개 = 총 {len(merged_data[component])}개 항목")
                            else:
                                # region 필드가 없는 경우 그대로 병합
                                merged_data[component] = df
                                print(f"  - {component}: region 필드 없음, 새 데이터 사용")
                        else:
                            # 기존 데이터가 없거나 비어있으면 새 데이터 사용
                            merged_data[component] = df
                            print(f"  - {component}: 새로운 데이터 {len(df)}개 항목 추가")
                
                # 시간 설정과 패턴 데이터 처리
                if 'timeseries' in new_data:
                    merged_data['timeseries'] = new_data['timeseries']
                
                if 'renewable_patterns' in new_data:
                    merged_data['renewable_patterns'] = new_data['renewable_patterns']
                
                if 'load_patterns' in new_data:
                    merged_data['load_patterns'] = new_data['load_patterns']
                
                if 'constraints' in new_data:
                    merged_data['constraints'] = new_data['constraints']
                
                # 통합 데이터 저장
                print("통합 데이터 생성 중...")
                with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
                    for component, df in merged_data.items():
                        if not df.empty:
                            df.to_excel(writer, sheet_name=component, index=False)
                
                print(f"통합 데이터가 '{self.output_path}'에 저장되었습니다.")
            
            return True
            
        except Exception as e:
            print(f"지역 데이터 저장 중 오류 발생: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def get_time_settings(self):
        """지역 데이터에서 사용된 시간 설정을 반환합니다."""
        try:
            # "시간 설정" 시트에서 직접 시간 설정 읽어오기
            if os.path.exists(self.template_path):
                try:
                    # 시간 설정 시트 읽기 시도
                    time_settings_df = pd.read_excel(self.template_path, sheet_name="시간 설정", engine="openpyxl")
                    
                    if not time_settings_df.empty:
                        # 첫 번째 행 가져오기
                        first_row = time_settings_df.iloc[0]
                        
                        # 시작 시간, 종료 시간, 간격 찾기
                        start_time = None
                        end_time = None
                        frequency = "1h"  # 기본값
                        
                        # 컬럼 이름으로 찾기
                        for col in time_settings_df.columns:
                            col_lower = col.lower()
                            if "시작" in col_lower or "start" in col_lower:
                                start_time = first_row[col]
                            elif "종료" in col_lower or "end" in col_lower:
                                end_time = first_row[col]
                            elif "간격" in col_lower or "주기" in col_lower or "freq" in col_lower:
                                frequency = first_row[col]
                        
                        # 시간 값이 찾아졌는지 확인
                        if start_time is not None and end_time is not None:
                            # datetime 객체로 변환
                            if isinstance(start_time, str):
                                start_time = pd.to_datetime(start_time)
                            if isinstance(end_time, str):
                                end_time = pd.to_datetime(end_time)
                                
                            # 문자열로 변환
                            start_time_str = start_time.strftime('%Y-%m-%d %H:%M:%S')
                            end_time_str = end_time.strftime('%Y-%m-%d %H:%M:%S')
                            
                            print(f"시간 설정 시트에서 시간 설정을 읽었습니다: {start_time_str} ~ {end_time_str}, 간격: {frequency}")
                            
                            return {
                                'start_time': start_time_str,
                                'end_time': end_time_str,
                                'frequency': str(frequency)
                            }
                except Exception as e:
                    print(f"시간 설정 시트 읽기 실패: {str(e)}")
            
            # 기존 지역 통합 과정에서 설정된 시간 설정 확인
            if hasattr(self.data_manager, 'timeseries') and self.data_manager.timeseries is not None:
                ts = self.data_manager.timeseries.iloc[0]
                if 'start_time' in ts and 'end_time' in ts and 'frequency' in ts:
                    start_time = ts['start_time']
                    end_time = ts['end_time']
                    frequency = ts['frequency']
                    
                    print(f"지역 데이터 관리자에서 시간 설정을 가져왔습니다: {start_time} ~ {end_time}, 간격: {frequency}")
                    
                    return {
                        'start_time': str(start_time),
                        'end_time': str(end_time),
                        'frequency': str(frequency)
                    }
            
            # 대체 기본값 반환
            print("시간 설정을 찾을 수 없어 기본값 사용: 2023-01-01 00:00:00 ~ 2023-12-30 00:00:00, 간격: 1h")
            return {
                'start_time': '2023-01-01 00:00:00',
                'end_time': '2023-12-30 00:00:00',
                'frequency': '1h'
            }
        except Exception as e:
            print(f"시간 설정 가져오기 실패: {str(e)}")
            return None
    
    def run_model_directly(self, selected_regions=None, connections=None):
        """직접 모델 실행 (프로그래매틱 인터페이스)
        
        Args:
            selected_regions (list): 선택할 지역 코드 목록 (예: ['SEL', 'JBD'])
            connections (list): 지역간 연결 목록 (예: [('SEL', 'JBD', {'s_nom': 1000})])
            
        Returns:
            bool: 실행 성공 여부
        """
        try:
            # 선택된 지역 설정
            if not selected_regions:
                print("선택된 지역이 없습니다. 기본 지역(서울, 전북)을 사용합니다.")
                selected_regions = ['SEL', 'JBD']
            
            # 지역 초기화
            for region_code in selected_regions:
                self.region_selector.select_region(region_code)
                self.data_manager.initialize_region(region_code)
            
            # 지역간 연결 추가
            if connections:
                for conn in connections:
                    region1, region2, conn_data = conn if len(conn) == 3 else (conn[0], conn[1], None)
                    self.data_manager.add_connection(region1, region2, conn_data)
            else:
                # 기본 연결 추가
                if 'SEL' in selected_regions and 'JBD' in selected_regions:
                    self.data_manager.add_connection('SEL', 'JBD')
            
            # 통합 데이터 생성 및 저장
            merged_data = self.data_manager.merge_data()
            if not merged_data:
                print("통합 데이터를 생성할 수 없습니다.")
                return False
            
            self.data_manager.export_merged_data(self.output_path)
            
            # PyPSA 모델 실행
            print("\nPyPSA 모델 실행 중...")
            
            # 명령어로 PyPSA_GUI 실행 (integrated_input_data.xlsx를 직접 사용)
            import subprocess
            cmd = ["python", "PyPSA_GUI.py", self.output_path]
            print(f"명령어 실행: {' '.join(cmd)}")
            
            result = subprocess.run(cmd)
            if result.returncode != 0:
                print("PyPSA_GUI 실행 실패")
                return False
            
            print(f"모델 실행이 완료되었습니다.")
            return True
            
        except Exception as e:
            print(f"모델 실행 중 오류 발생: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def show_regions(self):
        """사용 가능한 지역 목록 표시"""
        print("\n사용 가능한 지역 목록:")
        print("-" * 50)
        print(f"{'코드':<6} {'지역명':<15} {'인구(명)':<12} {'면적(km²)':<10}")
        print("-" * 50)
        
        for code, region in sorted(self.region_selector.regions.items()):
            print(f"{code:<6} {region['name']:<15} {region['population']:<12,} {region['area']:<10.1f}")
        
        print("-" * 50)
        return True
    
    def visualize_map(self, selected_regions=None):
        """지도 시각화
        
        Args:
            selected_regions (list): 강조 표시할 지역 코드 목록
            
        Returns:
            bool: 성공 여부
        """
        try:
            import matplotlib.pyplot as plt
            
            # 선택된 지역 설정
            if selected_regions:
                for region_code in selected_regions:
                    print(f"지역 선택: {self.region_selector.regions[region_code]['name']}")
                    self.region_selector.select_region(region_code)
            
            # 지도 그리기
            fig, ax = self.region_selector.draw_korea_map()
            plt.tight_layout()
            plt.savefig('korea_regions.png', dpi=300, bbox_inches='tight')
            plt.close()
            
            print("지도가 'korea_regions.png'에 저장되었습니다.")
            return True
            
        except Exception as e:
            print(f"지도 시각화 중 오류 발생: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def print_help(self):
        """도움말 출력"""
        print("\nPyPSA-HD 지역 기반 인터페이스 사용법")
        print("=" * 50)
        print("1. create_template: Excel 템플릿 생성")
        print("2. process_template: 입력된 템플릿 처리 및 모델 실행")
        print("3. save_region_data: 특정 지역 또는 모든 선택된 지역의 데이터 저장")
        print("4. run_model: 선택한 지역으로 모델 직접 실행")
        print("5. show_regions: 사용 가능한 지역 목록 표시")
        print("6. visualize_map: 지역 지도 시각화")
        print("7. help: 도움말 표시")
        print("=" * 50)
        print("예시:")
        print("python PyPSA_HD_Regional.py create_template")
        print("python PyPSA_HD_Regional.py process_template --input my_template.xlsx")
        print("python PyPSA_HD_Regional.py save_region_data --region SEL")
        print("python PyPSA_HD_Regional.py run_model --regions SEL,JBD,GGD")
        print("python PyPSA_HD_Regional.py show_regions")
        print("python PyPSA_HD_Regional.py visualize_map --regions SEL,JBD")
        print("=" * 50)

def main():
    """메인 함수"""
    # 인터페이스 인스턴스 생성
    interface = PyPSA_HD_Regional()
    
    # 명령행 인수 파싱
    parser = argparse.ArgumentParser(description='PyPSA-HD 지역 기반 인터페이스')
    parser.add_argument('command', choices=['create_template', 'process_template', 'save_region_data', 'run_model',
                                           'show_regions', 'visualize_map', 'help'],
                       help='실행할 명령')
    parser.add_argument('--input', type=str, help='입력 Excel 파일 경로')
    parser.add_argument('--output', type=str, help='출력 파일 경로')
    parser.add_argument('--regions', type=str, help='지역 코드 목록 (쉼표로 구분)')
    parser.add_argument('--region', type=str, help='단일 지역 코드')
    args = parser.parse_args()
    
    # 출력 파일 경로 설정
    if args.output:
        interface.output_path = args.output
    
    # 명령 실행
    if args.command == 'create_template':
        interface.create_template()
    
    elif args.command == 'process_template':
        interface.process_template(args.input)
    
    elif args.command == 'save_region_data':
        # 지역 데이터 저장
        interface.save_region_data(args.input, args.region)
    
    elif args.command == 'run_model':
        # 지역 목록 파싱
        regions = args.regions.split(',') if args.regions else None
        interface.run_model_directly(regions)
    
    elif args.command == 'show_regions':
        interface.show_regions()
    
    elif args.command == 'visualize_map':
        # 지역 목록 파싱
        regions = args.regions.split(',') if args.regions else None
        interface.visualize_map(regions)
    
    elif args.command == 'help':
        interface.print_help()

if __name__ == "__main__":
    main() 