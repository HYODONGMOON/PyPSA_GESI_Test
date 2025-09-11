import pypsa
import pandas as pd
import sys
import os
import traceback
import argparse
from datetime import datetime
import logging

# 로그 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("cplex_optim.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("PyPSA-HD-Optimizer")

def run_optimization(input_file, time_period='week'):
    """
    CPLEX 솔버를 사용하여 PyPSA 모델을 최적화합니다.
    
    Args:
        input_file: 입력 Excel 파일 경로
        time_period: 시간 범위 ('day', 'week', 'month')
    """
    logger.info(f"PyPSA-HD 모델 최적화 시작 (파일: {input_file}, 기간: {time_period})")
    logger.info(f"Python 버전: {sys.version}")
    logger.info(f"PyPSA 버전: {pypsa.__version__}")
    
    # 시간 범위 설정
    if time_period == 'day':
        snapshots = pd.date_range("2023-01-01 00:00", "2023-01-01 23:00", freq="1h")
        logger.info("1일 기간으로 최적화합니다.")
    elif time_period == 'week':
        snapshots = pd.date_range("2023-01-01 00:00", "2023-01-07 23:00", freq="1h")
        logger.info("1주일 기간으로 최적화합니다.")
    elif time_period == 'month':
        snapshots = pd.date_range("2023-01-01 00:00", "2023-01-31 23:00", freq="1h")
        logger.info("1개월 기간으로 최적화합니다.")
    else:
        logger.error(f"잘못된 시간 범위: {time_period}")
        return False
    
    try:
        # 네트워크 생성
        network = pypsa.Network()
        
        # 스냅샷 설정
        network.set_snapshots(snapshots)
        logger.info(f"스냅샷이 설정되었습니다: {len(network.snapshots)}개")
        
        # 데이터 로드
        logger.info(f"Excel 파일 '{input_file}'에서 데이터 로드 중...")
        
        with pd.ExcelFile(input_file) as xls:
            available_sheets = xls.sheet_names
            logger.info(f"사용 가능한 시트: {available_sheets}")
        
        # 버스 데이터 로드
        if 'buses' in available_sheets:
            buses = pd.read_excel(input_file, sheet_name='buses')
            for i, row in buses.iterrows():
                network.add("Bus", 
                           row['name'],
                           v_nom=row['v_nom'] if 'v_nom' in buses.columns and not pd.isna(row['v_nom']) else 380,
                           carrier=row['carrier'] if 'carrier' in buses.columns and not pd.isna(row['carrier']) else 'AC',
                           x=row['x'] if 'x' in buses.columns and not pd.isna(row['x']) else 0,
                           y=row['y'] if 'y' in buses.columns and not pd.isna(row['y']) else 0)
            logger.info(f"버스 데이터가 로드되었습니다: {len(buses)}개")
        else:
            logger.error("'buses' 시트를 찾을 수 없습니다.")
            return False
        
        # 발전기 데이터 로드
        if 'generators' in available_sheets:
            generators = pd.read_excel(input_file, sheet_name='generators')
            for i, row in generators.iterrows():
                network.add("Generator",
                           row['name'],
                           bus=row['bus'],
                           p_nom=row['p_nom'] if 'p_nom' in generators.columns and not pd.isna(row['p_nom']) else 0,
                           p_nom_extendable=row['p_nom_extendable'] if 'p_nom_extendable' in generators.columns and not pd.isna(row['p_nom_extendable']) else True,
                           p_nom_min=row['p_nom_min'] if 'p_nom_min' in generators.columns and not pd.isna(row['p_nom_min']) else 0,
                           p_nom_max=row['p_nom_max'] if 'p_nom_max' in generators.columns and not pd.isna(row['p_nom_max']) else 10000,
                           marginal_cost=row['marginal_cost'] if 'marginal_cost' in generators.columns and not pd.isna(row['marginal_cost']) else 50,
                           carrier=row['carrier'] if 'carrier' in generators.columns and not pd.isna(row['carrier']) else 'AC')
            logger.info(f"발전기 데이터가 로드되었습니다: {len(generators)}개")
        else:
            logger.warning("'generators' 시트를 찾을 수 없습니다.")
        
        # 부하 데이터 로드
        if 'loads' in available_sheets:
            loads = pd.read_excel(input_file, sheet_name='loads')
            for i, row in loads.iterrows():
                # 모든 스냅샷에 대해 단일 값 설정
                p_set = pd.Series(
                    data=row['p_set'] if 'p_set' in loads.columns and not pd.isna(row['p_set']) else 0,
                    index=network.snapshots
                )
                network.add("Load",
                           row['name'],
                           bus=row['bus'],
                           p_set=p_set)
            logger.info(f"부하 데이터가 로드되었습니다: {len(loads)}개")
        else:
            logger.warning("'loads' 시트를 찾을 수 없습니다.")
        
        # 선로 데이터 로드
        if 'lines' in available_sheets:
            lines = pd.read_excel(input_file, sheet_name='lines')
            for i, row in lines.iterrows():
                network.add("Line",
                           row['name'],
                           bus0=row['bus0'],
                           bus1=row['bus1'],
                           x=row['x'] if 'x' in lines.columns and not pd.isna(row['x']) else 0.1,
                           r=row['r'] if 'r' in lines.columns and not pd.isna(row['r']) else 0.01,
                           s_nom=row['s_nom'] if 's_nom' in lines.columns and not pd.isna(row['s_nom']) else 1000,
                           s_nom_extendable=row['s_nom_extendable'] if 's_nom_extendable' in lines.columns and not pd.isna(row['s_nom_extendable']) else True,
                           s_nom_min=row['s_nom_min'] if 's_nom_min' in lines.columns and not pd.isna(row['s_nom_min']) else 0,
                           s_nom_max=row['s_nom_max'] if 's_nom_max' in lines.columns and not pd.isna(row['s_nom_max']) else 10000,
                           carrier=row['carrier'] if 'carrier' in lines.columns and not pd.isna(row['carrier']) else 'AC')
            logger.info(f"선로 데이터가 로드되었습니다: {len(lines)}개")
        else:
            logger.warning("'lines' 시트를 찾을 수 없습니다.")
        
        # 링크 데이터 로드
        if 'links' in available_sheets:
            links = pd.read_excel(input_file, sheet_name='links')
            for i, row in links.iterrows():
                network.add("Link",
                           row['name'],
                           bus0=row['bus0'],
                           bus1=row['bus1'],
                           p_nom=row['p_nom'] if 'p_nom' in links.columns and not pd.isna(row['p_nom']) else 1000,
                           p_nom_extendable=row['p_nom_extendable'] if 'p_nom_extendable' in links.columns and not pd.isna(row['p_nom_extendable']) else True,
                           p_nom_min=row['p_nom_min'] if 'p_nom_min' in links.columns and not pd.isna(row['p_nom_min']) else 0,
                           p_nom_max=row['p_nom_max'] if 'p_nom_max' in links.columns and not pd.isna(row['p_nom_max']) else 10000,
                           efficiency=row['efficiency'] if 'efficiency' in links.columns and not pd.isna(row['efficiency']) else 1.0,
                           marginal_cost=row['marginal_cost'] if 'marginal_cost' in links.columns and not pd.isna(row['marginal_cost']) else 0)
            logger.info(f"링크 데이터가 로드되었습니다: {len(links)}개")
        else:
            logger.warning("'links' 시트를 찾을 수 없습니다.")
        
        # 저장장치 데이터 로드
        if 'stores' in available_sheets:
            stores = pd.read_excel(input_file, sheet_name='stores')
            for i, row in stores.iterrows():
                network.add("Store",
                           row['name'],
                           bus=row['bus'],
                           e_nom=row['e_nom'] if 'e_nom' in stores.columns and not pd.isna(row['e_nom']) else 1000,
                           e_nom_extendable=row['e_nom_extendable'] if 'e_nom_extendable' in stores.columns and not pd.isna(row['e_nom_extendable']) else True,
                           e_nom_max=row['e_nom_max'] if 'e_nom_max' in stores.columns and not pd.isna(row['e_nom_max']) else 10000,
                           e_cyclic=row['e_cyclic'] if 'e_cyclic' in stores.columns and not pd.isna(row['e_cyclic']) else True,
                           standing_loss=row['standing_loss'] if 'standing_loss' in stores.columns and not pd.isna(row['standing_loss']) else 0.0,
                           carrier=row['carrier'] if 'carrier' in stores.columns and not pd.isna(row['carrier']) else 'AC')
            logger.info(f"저장장치 데이터가 로드되었습니다: {len(stores)}개")
        else:
            logger.warning("'stores' 시트를 찾을 수 없습니다.")
        
        # 네트워크 요약
        logger.info("\n네트워크 구성 요약:")
        logger.info(f"버스: {len(network.buses)}개")
        logger.info(f"발전기: {len(network.generators)}개")
        logger.info(f"부하: {len(network.loads)}개")
        logger.info(f"선로: {len(network.lines)}개")
        logger.info(f"링크: {len(network.links)}개")
        logger.info(f"저장장치: {len(network.stores)}개")
        logger.info(f"스냅샷: {len(network.snapshots)}개")
        
        # 최적화 시작
        logger.info("\n최적화 시작...")
        
        # 최적화 이전 네트워크 저장
        try:
            network.export_to_netcdf(f"{os.path.splitext(input_file)[0]}_before_optim.nc")
            logger.info(f"최적화 전 네트워크가 저장되었습니다.")
        except Exception as e:
            logger.warning(f"최적화 전 네트워크 저장 실패: {e}")
        
        # CPLEX 솔버 설정
        solver_options = {
            "threads": 4,
            "lpmethod": 4,  # 배리어 메서드
            "barrier.algorithm": 3  # 기본 배리어 알고리즘
            # solutiontype 파라미터 제거
        }
        
        logger.info("CPLEX 솔버로 최적화 시도 중...")
        
        try:
            # CPLEX로 최적화
            network.optimize(solver_name="cplex", solver_options=solver_options)
            logger.info("CPLEX 최적화 성공!")
            
            # 결과 저장
            try:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_file = f"{os.path.splitext(input_file)[0]}_results_{timestamp}.nc"
                network.export_to_netcdf(output_file)
                logger.info(f"최적화 결과가 '{output_file}'에 저장되었습니다.")
                
                # 결과를 Excel로도 저장
                excel_output = f"{os.path.splitext(input_file)[0]}_results_{timestamp}.xlsx"
                with pd.ExcelWriter(excel_output) as writer:
                    # 목적 함수 값
                    pd.DataFrame({'objective': [network.objective]}).to_excel(writer, sheet_name='summary', index=False)
                    
                    # 발전기 결과
                    if not network.generators_t.p.empty:
                        network.generators_t.p.to_excel(writer, sheet_name='generators_p')
                    
                    # 선로 결과
                    if not network.lines_t.p0.empty:
                        network.lines_t.p0.to_excel(writer, sheet_name='lines_p0')
                    
                    # 링크 결과
                    if not network.links_t.p0.empty:
                        network.links_t.p0.to_excel(writer, sheet_name='links_p0')
                    
                    # 저장장치 결과
                    if hasattr(network, 'stores_t') and hasattr(network.stores_t, 'p') and not network.stores_t.p.empty:
                        network.stores_t.p.to_excel(writer, sheet_name='stores_p')
                
                logger.info(f"최적화 결과가 Excel 파일 '{excel_output}'에도 저장되었습니다.")
                
                # 결과 요약
                logger.info("\n최적화 결과 요약:")
                logger.info(f"목적 함수 값: {network.objective:.2f}")
                
                return True
            except Exception as e:
                logger.error(f"결과 저장 중 오류 발생: {e}")
                traceback.print_exc()
        except Exception as e:
            logger.error(f"CPLEX 최적화 오류: {e}")
            traceback.print_exc()
            
            # 다른 솔버 시도
            try:
                logger.info("\nGLPK 솔버로 다시 시도 중...")
                network.optimize(solver_name="glpk")
                logger.info("GLPK 최적화 성공!")
                
                # 결과 저장
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_file = f"{os.path.splitext(input_file)[0]}_results_glpk_{timestamp}.nc"
                network.export_to_netcdf(output_file)
                logger.info(f"GLPK 최적화 결과가 '{output_file}'에 저장되었습니다.")
                
                return True
            except Exception as e2:
                logger.error(f"GLPK 솔버 오류: {e2}")
                traceback.print_exc()
                return False
    
    except Exception as e:
        logger.error(f"모델 생성 중 오류 발생: {e}")
        traceback.print_exc()
        return False

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='CPLEX 솔버를 사용하여 PyPSA 모델을 최적화합니다.')
    parser.add_argument('--input', default='simplified_input_data.xlsx', 
                        help='입력 파일 경로 (기본값: simplified_input_data.xlsx)')
    parser.add_argument('--time', choices=['day', 'week', 'month'], default='week',
                        help='최적화 기간 (기본값: week)')
    args = parser.parse_args()
    
    result = run_optimization(args.input, args.time)
    
    status = "성공" if result else "실패"
    logger.info(f"\n최종 실행 결과: {status}")
    logger.info(f"로그 파일이 'cplex_optim.log'에 저장되었습니다.") 