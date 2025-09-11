import pandas as pd
import os
import numpy as np
import glob
from datetime import datetime

# analyze_regional_results.py에서 함수를 재사용합니다
def get_latest_timestamp():
    """최신 타임스탬프 가져오기"""
    timestamp_file = os.path.join('results', "latest_timestamp.txt")
    if os.path.exists(timestamp_file):
        with open(timestamp_file, 'r') as f:
            timestamp = f.read().strip()
            return timestamp
    
    # 파일이 없는 경우 최신 결과 폴더 검색
    result_dirs = glob.glob(os.path.join('results', "[0-9]" * 8 + "_" + "[0-9]" * 6))
    if result_dirs:
        # 폴더 수정 시간으로 정렬
        latest_dir = max(result_dirs, key=os.path.getmtime)
        return os.path.basename(latest_dir)
    
    # 아무것도 없는 경우 예외 발생
    raise FileNotFoundError("최신 타임스탬프 정보를 찾을 수 없습니다.")

def get_results_dir():
    """최신 결과 디렉토리 경로 가져오기"""
    try:
        timestamp = get_latest_timestamp()
        results_dir = os.path.join('results', timestamp)
        if os.path.exists(results_dir):
            return results_dir
    except Exception as e:
        print(f"최신 결과 디렉토리를 찾는 중 오류 발생: {str(e)}")
    
    # 실패한 경우 기본 폴더 반환
    return 'results'

# 최신 타임스탬프 폴더 가져오기
try:
    results_dir = get_results_dir()
    print(f"분석 결과를 '{results_dir}' 폴더에 저장합니다.")
except Exception as e:
    print(f"타임스탬프 폴더를 찾는 데 실패했습니다: {str(e)}")
    results_dir = 'results'
    print(f"기본 결과 폴더 '{results_dir}'를 사용합니다.")

# 결과를 저장할 Excel 파일 경로
output_file = os.path.join(results_dir, 'combined_results.xlsx')

print("파일 통합 시작...")
print("주의: 파일 크기가 매우 클 수 있으므로 시간이 오래 걸릴 수 있습니다.")

# 최신 타임스탬프 (가장 최근 결과 파일의 타임스탬프)
latest_timestamp = get_latest_timestamp()

# Excel 작성기 생성
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    
    # 1. 최신 결과 파일 처리 (results 폴더에서)
    # 1-1. 발전기 출력 데이터
    generator_file = os.path.join('results', latest_timestamp, f'optimization_result_{latest_timestamp}_generator_output.csv')
    if not os.path.exists(generator_file):
        generator_file = os.path.join('results', f'optimization_result_{latest_timestamp}_generator_output.csv')
    
    if os.path.exists(generator_file):
        print(f"\n파일 처리 중: {generator_file}")
        try:
            # 전체 데이터 로드 (시간이 오래 걸릴 수 있음)
            df = pd.read_csv(generator_file)
            df.to_excel(writer, sheet_name='Generator_Output', index=False)
            print(f"  - 발전기 출력 데이터 처리 완료 (행 수: {len(df)})")
        except Exception as e:
            print(f"  - 오류 발생: {str(e)}")
    
    # 1-2. 부하 데이터
    load_file = os.path.join('results', latest_timestamp, f'optimization_result_{latest_timestamp}_load.csv')
    if not os.path.exists(load_file):
        load_file = os.path.join('results', f'optimization_result_{latest_timestamp}_load.csv')
    
    if os.path.exists(load_file):
        print(f"\n파일 처리 중: {load_file}")
        try:
            # 전체 데이터 로드
            df = pd.read_csv(load_file)
            df.to_excel(writer, sheet_name='Load_Data', index=False)
            print(f"  - 부하 데이터 처리 완료 (행 수: {len(df)})")
        except Exception as e:
            print(f"  - 오류 발생: {str(e)}")
    
    # 1-3. 라인 사용률 데이터
    line_file = os.path.join('results', latest_timestamp, f'optimization_result_{latest_timestamp}_line_usage.csv')
    if not os.path.exists(line_file):
        line_file = os.path.join('results', f'optimization_result_{latest_timestamp}_line_usage.csv')
    
    if os.path.exists(line_file):
        print(f"\n파일 처리 중: {line_file}")
        try:
            # 전체 데이터 로드
            df = pd.read_csv(line_file)
            df.to_excel(writer, sheet_name='Line_Usage', index=False)
            print(f"  - 라인 사용률 데이터 처리 완료 (행 수: {len(df)})")
        except Exception as e:
            print(f"  - 오류 발생: {str(e)}")
    
    # 1-4. 저장장치 데이터
    storage_file = os.path.join('results', latest_timestamp, f'optimization_result_{latest_timestamp}_storage.csv')
    if not os.path.exists(storage_file):
        storage_file = os.path.join('results', f'optimization_result_{latest_timestamp}_storage.csv')
    
    if os.path.exists(storage_file):
        print(f"\n파일 처리 중: {storage_file}")
        try:
            # 전체 데이터 로드
            df = pd.read_csv(storage_file)
            df.to_excel(writer, sheet_name='Storage_Data', index=False)
            print(f"  - 저장장치 데이터 처리 완료 (행 수: {len(df)})")
        except Exception as e:
            print(f"  - 오류 발생: {str(e)}")
            
    # 2. 지역별 에너지 균형 데이터 추가
    balance_file = os.path.join(results_dir, 'regional_energy_balance.csv')
    
    if os.path.exists(balance_file):
        print(f"\n파일 처리 중: {balance_file}")
        try:
            df = pd.read_csv(balance_file)
            df.to_excel(writer, sheet_name='Regional_Energy_Balance', index=False)
            print(f"  - 지역별 에너지 균형 데이터 처리 완료 (행 수: {len(df)})")
        except Exception as e:
            print(f"  - 오류 발생: {str(e)}")
    
    # 3. 송전망 흐름 데이터 추가
    flow_file = os.path.join(results_dir, 'transmission_flow.csv')
    
    if os.path.exists(flow_file):
        print(f"\n파일 처리 중: {flow_file}")
        try:
            df = pd.read_csv(flow_file)
            df.to_excel(writer, sheet_name='Transmission_Flow', index=False)
            print(f"  - 송전망 흐름 데이터 처리 완료 (행 수: {len(df)})")
        except Exception as e:
            print(f"  - 오류 발생: {str(e)}")
    
    # 4. 분석 결과 파일들 처리
    # 최신 타임스탬프 사용
    analysis_timestamp = latest_timestamp
    
    # 각 분석 결과에 대한 파일 경로 패턴을 정의합니다
    analysis_file_patterns = [
        {'name': '지역별_발전원별_발전량', 'sheet': '지역별_발전원별_발전량'},
        {'name': '지역별_발전량', 'sheet': '지역별_발전량'},
        {'name': '발전원별_발전량', 'sheet': '발전원별_발전량'},
        {'name': '상위발전기_발전량', 'sheet': '상위발전기_발전량'}
    ]
    
    # 각 분석 결과 파일에 대해 처리
    for pattern in analysis_file_patterns:
        file_name = pattern['name']
        sheet_name = pattern['sheet']
        
        # 타임스탬프 폴더 내의 파일 경로
        file_path = os.path.join(results_dir, f'optimization_result_{analysis_timestamp}_{file_name}.csv')
        
        if os.path.exists(file_path):
            print(f"\n파일 처리 중: {file_path}")
            try:
                # 데이터 로드 및 Excel에 저장
                df = pd.read_csv(file_path)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"  - {sheet_name} 데이터 처리 완료 (행 수: {len(df)})")
            except Exception as e:
                print(f"  - 오류 발생: {str(e)}")
        else:
            print(f"\n  - '{file_name}' 분석 결과 파일을 찾을 수 없습니다.")

print(f"\n모든 데이터가 '{output_file}'에 통합되었습니다.")
print(f"결과 파일이 '{results_dir}' 폴더에 저장되었습니다.") 