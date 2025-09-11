import pypsa
import pandas as pd
import sys
import os
import traceback
import logging
from datetime import datetime, timedelta

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("pypsa_cplex_log.txt"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("PyPSA-HD-Optimizer")

def fix_and_optimize(input_file='integrated_input_data.xlsx', time_period='day'):
    """
    PyPSA-HD 모델의 최적화 수행
    
    Args:
        input_file: 입력 데이터가 있는 Excel 파일 경로
        time_period: 'day', 'week', 'month' 중 하나 선택
    
    Returns:
        성공 여부(True/False)
    """
    logger.info(f"PyPSA-HD 모델 최적화 시작")
    logger.info(f"Python 버전: {sys.version}")
    logger.info(f"PyPSA 버전: {pypsa.__version__}")
    
    # 시간 범위 설정
    if time_period == 'day':
        start_time = '2023-01-01 00:00:00'
        end_time = '2023-01-02 00:00:00'  # 1일
        logger.info("시간 범위: 1일 (24시간)")
    elif time_period == 'week':
        start_time = '2023-01-01 00:00:00'
        end_time = '2023-01-08 00:00:00'  # 1주일
        logger.info("시간 범위: 1주일 (168시간)")
    elif time_period == 'month':
        start_time = '2023-01-01 00:00:00'
        end_time = '2023-02-01 00:00:00'  # 1개월
        logger.info("시간 범위: 1개월 (744시간)")
    else:
        logger.error(f"잘못된 시간 범위 옵션: {time_period} (day, week, month 중 하나여야 함)")
        return False
    
    try:
        # 입력 파일 존재 여부 확인
        if not os.path.exists(input_file):
            logger.error(f"입력 파일 '{input_file}'이 존재하지 않습니다.")
            return False
        
        # 네트워크 생성
        network = pypsa.Network()
        
        # 시간 범위 설정
        snapshots = pd.date_range(
            start=start_time, 
            end=end_time, 
            freq='1h',
            inclusive='left'
        )
        network.set_snapshots(snapshots)
        
        # 데이터 로드
        logger.info(f"'{input_file}'에서 데이터 로드 중...")
        
        # 각 시트 로드
        sheets_to_check = ['buses', 'generators', 'loads', 'lines', 'links', 'stores']
        loaded_components = {}
        
        with pd.ExcelFile(input_file) as xls:
            for sheet in sheets_to_check:
                if sheet in xls.sheet_names:
                    loaded_components[sheet] = pd.read_excel(xls, sheet_name=sheet)
                    logger.info(f"'{sheet}' 시트에서 {len(loaded_components[sheet])}개 행 로드됨")
                else:
                    logger.warning(f"'{sheet}' 시트를 찾을 수 없습니다.")
                    loaded_components[sheet] = pd.DataFrame()
        
        # 기본 체크
        for sheet in ['buses', 'generators', 'loads']:
            if sheet not in loaded_components or loaded_components[sheet].empty:
                logger.error(f"필수 시트 '{sheet}'가 비어 있거나 찾을 수 없습니다.")
                return False
        
        # carrier 속성 확인 및 수정
        for sheet in sheets_to_check:
            if sheet in loaded_components and not loaded_components[sheet].empty:
                if 'carrier' in loaded_components[sheet].columns:
                    null_carriers = loaded_components[sheet]['carrier'].isnull().sum()
                    if null_carriers > 0:
                        logger.warning(f"'{sheet}' 시트에서 {null_carriers}개의 null carrier 값을 기본값으로 대체합니다.")
                        
                        # 기본값 설정
                        if sheet == 'buses' or sheet == 'generators' or sheet == 'lines':
                            loaded_components[sheet]['carrier'] = loaded_components[sheet]['carrier'].fillna('AC')
                        elif sheet == 'links':
                            loaded_components[sheet]['carrier'] = loaded_components[sheet]['carrier'].fillna('Link')
                        elif sheet == 'stores':
                            loaded_components[sheet]['carrier'] = loaded_components[sheet]['carrier'].fillna('Store')
        
        # 네트워크에 컴포넌트 추가
        available_buses = set()
        
        # 버스 추가
        for _, bus in loaded_components['buses'].iterrows():
            bus_name = str(bus['name'])
            network.add("Bus",
                     name=bus_name,
                     v_nom=float(bus['v_nom']),
                     carrier=str(bus['carrier'] if pd.notna(bus['carrier']) else 'AC'))
            available_buses.add(bus_name)
        
        # 발전기 추가
        for _, gen in loaded_components['generators'].iterrows():
            bus_name = str(gen['bus'])
            if bus_name not in available_buses:
                logger.warning(f"발전기 '{gen['name']}'의 버스 '{bus_name}'가 정의되지 않았습니다. 해당 발전기를 건너뜁니다.")
                continue
            
            params = {
                'name': str(gen['name']),
                'bus': bus_name,
                'p_nom': float(gen['p_nom']),
                'carrier': str(gen['carrier'] if pd.notna(gen['carrier']) else 'AC')
            }
            
            # 선택적 매개변수 설정
            for param in ['p_nom_extendable', 'p_nom_min', 'p_nom_max', 'marginal_cost', 'capital_cost']:
                if param in gen and pd.notna(gen[param]):
                    params[param] = float(gen[param])
            
            # p_nom_extendable이 True인 경우, p_nom_max가 없으면 매우 큰 값으로 설정
            if 'p_nom_extendable' in params and params['p_nom_extendable'] and ('p_nom_max' not in params or pd.isna(params['p_nom_max'])):
                params['p_nom_max'] = 10000  # 매우 큰 값
            
            network.add("Generator", **params)
        
        # 부하 추가
        for _, load in loaded_components['loads'].iterrows():
            bus_name = str(load['bus'])
            if bus_name not in available_buses:
                logger.warning(f"부하 '{load['name']}'의 버스 '{bus_name}'가 정의되지 않았습니다. 해당 부하를 건너뜁니다.")
                continue
            
            network.add("Load",
                      name=str(load['name']),
                      bus=bus_name,
                      p_set=float(load['p_set']))
        
        # 선로 추가
        if not loaded_components['lines'].empty:
            for _, line in loaded_components['lines'].iterrows():
                bus0_name = str(line['bus0'])
                bus1_name = str(line['bus1'])
                
                if bus0_name not in available_buses:
                    logger.warning(f"선로 '{line['name']}'의 bus0 '{bus0_name}'가 정의되지 않았습니다. 해당 선로를 건너뜁니다.")
                    continue
                if bus1_name not in available_buses:
                    logger.warning(f"선로 '{line['name']}'의 bus1 '{bus1_name}'가 정의되지 않았습니다. 해당 선로를 건너뜁니다.")
                    continue
                
                params = {
                    'name': str(line['name']),
                    'bus0': bus0_name,
                    'bus1': bus1_name,
                    's_nom': float(line['s_nom']),
                    'carrier': str(line['carrier'] if pd.notna(line['carrier']) else 'AC')
                }
                
                for param in ['x', 'r', 's_nom_extendable', 's_nom_max']:
                    if param in line and pd.notna(line[param]):
                        params[param] = float(line[param])
                
                # s_nom_extendable이 True인 경우, s_nom_max가 없으면 매우 큰 값으로 설정
                if 's_nom_extendable' in params and params['s_nom_extendable'] and ('s_nom_max' not in params or pd.isna(params['s_nom_max'])):
                    params['s_nom_max'] = 10000  # 매우 큰 값
                
                # x, r 값이 누락된 경우 기본값 설정
                if 'x' not in params or pd.isna(params['x']):
                    params['x'] = 0.1  # 기본 리액턴스 값
                if 'r' not in params or pd.isna(params['r']):
                    params['r'] = 0.01  # 기본 저항 값
                
                network.add("Line", **params)
        
        # 링크 추가
        if not loaded_components['links'].empty:
            for _, link in loaded_components['links'].iterrows():
                bus0_name = str(link['bus0'])
                
                if bus0_name not in available_buses:
                    logger.warning(f"링크 '{link['name']}'의 bus0 '{bus0_name}'가 정의되지 않았습니다. 해당 링크를 건너뜁니다.")
                    continue
                
                params = {
                    'name': str(link['name']),
                    'bus0': bus0_name,
                    'p_nom': float(link['p_nom']),
                    'carrier': str(link['carrier'] if pd.notna(link['carrier']) else 'Link')
                }
                
                # bus1 확인 및 추가
                if 'bus1' in link and pd.notna(link['bus1']):
                    bus1_name = str(link['bus1'])
                    if bus1_name not in available_buses:
                        logger.warning(f"링크 '{link['name']}'의 bus1 '{bus1_name}'가 정의되지 않았습니다. 해당 링크를 건너뜁니다.")
                        continue
                    params['bus1'] = bus1_name
                
                # 선택적 매개변수 추가
                for param in ['p_nom_extendable', 'p_nom_min', 'p_nom_max', 'efficiency', 'marginal_cost', 'capital_cost']:
                    if param in link and pd.notna(link[param]):
                        params[param] = float(link[param])
                
                # p_nom_extendable이 True인 경우, p_nom_max가 없으면 매우 큰 값으로 설정
                if 'p_nom_extendable' in params and params['p_nom_extendable'] and ('p_nom_max' not in params or pd.isna(params['p_nom_max'])):
                    params['p_nom_max'] = 10000  # 매우 큰 값
                
                network.add("Link", **params)
        
        # 저장장치 추가
        if not loaded_components['stores'].empty:
            for _, store in loaded_components['stores'].iterrows():
                bus_name = str(store['bus'])
                if bus_name not in available_buses:
                    logger.warning(f"저장장치 '{store['name']}'의 버스 '{bus_name}'가 정의되지 않았습니다. 해당 저장장치를 건너뜁니다.")
                    continue
                
                params = {
                    'name': str(store['name']),
                    'bus': bus_name,
                    'e_nom': float(store['e_nom']),
                    'carrier': str(store['carrier'] if pd.notna(store['carrier']) else 'Store')
                }
                
                # 선택적 매개변수 추가
                for param in ['e_nom_extendable', 'e_nom_max', 'e_cyclic', 'standing_loss', 'efficiency_store', 'efficiency_dispatch']:
                    if param in store and pd.notna(store[param]):
                        params[param] = float(store[param]) if param not in ['e_cyclic'] else bool(store[param])
                
                # e_nom_extendable이 True인 경우, e_nom_max가 없으면 매우 큰 값으로 설정
                if 'e_nom_extendable' in params and params['e_nom_extendable'] and ('e_nom_max' not in params or pd.isna(params['e_nom_max'])):
                    params['e_nom_max'] = 10000  # 매우 큰 값
                
                network.add("Store", **params)
        
        # 네트워크 요약 출력
        logger.info(f"네트워크 구성: {len(network.buses)} 버스, {len(network.generators)} 발전기, {len(network.loads)} 부하, {len(network.lines)} 선로, {len(network.links)} 링크, {len(network.stores)} 저장장치")
        
        # 외부 그리드 확인
        if 'External_Grid' not in network.buses.index:
            logger.warning("외부 그리드 버스가 없습니다. 외부 수입/수출이 제한될 수 있습니다.")
        
        # 백업 저장
        result_file = f'network_pre_optimization_{time_period}.nc'
        network.export_to_netcdf(result_file)
        logger.info(f"최적화 전 네트워크가 {result_file}에 저장되었습니다.")
        
        # 최적화 실행
        logger.info("CPLEX 최적화 시작...")
        
        # CPLEX 옵션 설정 - solutiontype 옵션 제거
        solver_options = {
            "threads": 4,
            "lpmethod": 4,  # 배리어 메서드
            "barrier.algorithm": 3,  # 기본 배리어 알고리즘
            "barrier.convergetol": 1e-4,  # 수렴 허용 오차 조절 (더 느슨하게)
            "barrier.limits.objrange": 1e10,  # objective range 제한 추가
            "optimalitytarget": 1  # 최적해만 찾기
        }
        
        try:
            # CPLEX로 최적화
            status = network.optimize(solver_name='cplex', solver_options=solver_options)
            logger.info(f"CPLEX 최적화 성공!")
            logger.info(f"목적 함수 값: {network.objective:.2f}")
            
            # 결과 저장 (NetCDF)
            result_file = f'network_results_{time_period}.nc'
            network.export_to_netcdf(result_file)
            logger.info(f"최적화 결과가 {result_file}에 저장되었습니다.")
            
            # 결과 저장 (Excel)
            result_file = f'results_{time_period}.xlsx'
            with pd.ExcelWriter(result_file) as writer:
                # 목적 함수 값 저장
                pd.DataFrame({
                    'Objective': [network.objective],
                    'Status': ['optimal'],
                    'Solver': ['cplex']
                }).to_excel(writer, sheet_name='Objective', index=False)
                
                # 발전기 결과
                if hasattr(network, 'generators_t') and hasattr(network.generators_t, 'p') and not network.generators_t.p.empty:
                    network.generators_t.p.to_excel(writer, sheet_name='Generator_Output')
                    
                    # 발전기 총합 계산
                    gen_sum = network.generators_t.p.sum(axis=1)
                    gen_sum.name = 'Total_Generation'
                    gen_sum.to_excel(writer, sheet_name='Generation_Summary')
                
                # 선로 결과
                if hasattr(network, 'lines_t') and hasattr(network.lines_t, 'p0') and not network.lines_t.p0.empty:
                    network.lines_t.p0.to_excel(writer, sheet_name='Line_Flows')
                
                # 링크 결과
                if hasattr(network, 'links_t') and hasattr(network.links_t, 'p0') and not network.links_t.p0.empty:
                    network.links_t.p0.to_excel(writer, sheet_name='Link_Flows')
                    
                    # 특히 Import/Export 링크 분석
                    import_export_links = [link for link in network.links_t.p0.columns if 'Import' in link or 'Export' in link]
                    if import_export_links:
                        network.links_t.p0[import_export_links].to_excel(writer, sheet_name='Import_Export_Flows')
                
                # 저장장치 결과
                if hasattr(network, 'stores_t') and hasattr(network.stores_t, 'e') and not network.stores_t.e.empty:
                    network.stores_t.e.to_excel(writer, sheet_name='Storage_Energy')
                
            logger.info(f"요약 결과가 {result_file}에 저장되었습니다.")
            
            return True
            
        except Exception as e:
            logger.error(f"CPLEX 최적화 중 오류 발생: {str(e)}")
            logger.error(traceback.format_exc())
            
            # GLPK로 재시도 (더 단순한 솔버)
            try:
                logger.info("GLPK로 재시도 중...")
                status = network.optimize(solver_name='glpk')
                logger.info(f"GLPK 최적화 성공!")
                logger.info(f"목적 함수 값: {network.objective:.2f}")
                
                # 결과 저장
                result_file = f'network_results_{time_period}_glpk.nc'
                network.export_to_netcdf(result_file)
                logger.info(f"GLPK 최적화 결과가 {result_file}에 저장되었습니다.")
                
                return True
                
            except Exception as glpk_e:
                logger.error(f"GLPK 최적화 중 오류 발생: {str(glpk_e)}")
                logger.error(traceback.format_exc())
                return False
    
    except Exception as e:
        logger.error(f"오류 발생: {str(e)}")
        logger.error(traceback.format_exc())
        return False

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description='PyPSA-HD 모델 최적화')
    parser.add_argument('--input', default='integrated_input_data.xlsx', 
                        help='입력 파일 경로 (기본값: integrated_input_data.xlsx)')
    parser.add_argument('--time', default='day', choices=['day', 'week', 'month'],
                        help='최적화 기간 (기본값: day)')
    
    args = parser.parse_args()
    
    success = fix_and_optimize(args.input, args.time)
    logger.info(f"최종 실행 결과: {'성공' if success else '실패'}") 