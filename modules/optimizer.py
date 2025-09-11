#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
최적화 모듈

PyPSA 네트워크 모델의 최적화를 수행합니다.
"""

import logging
import traceback
import multiprocessing
import time
import pypsa

logger = logging.getLogger("PyPSA-HD.Optimizer")

class PypsaOptimizer:
    """PyPSA 네트워크 최적화 클래스"""
    
    def __init__(self, config):
        """초기화 함수
        
        Args:
            config (dict): 설정 정보
        """
        self.config = config
        self.solver_name = config['solver']['name'] if 'solver' in config and 'name' in config['solver'] else 'cplex'
        self.solver_options = config['solver']['options'] if 'solver' in config and 'options' in config['solver'] else {}
        self.num_cores = config['solver']['threads'] if 'solver' in config and 'threads' in config['solver'] else multiprocessing.cpu_count()
    
    def optimize(self, network):
        """네트워크 최적화 실행
        
        Args:
            network (pypsa.Network): 최적화할 PyPSA 네트워크 객체
            
        Returns:
            bool: 최적화 성공 여부
        """
        if network is None:
            logger.error("네트워크가 생성되지 않았습니다.")
            return False
        
        try:
            # 사용 가능한 솔버 확인 및 선택
            solver_name = self.solver_name
            
            # 솔버 사용 가능 여부 확인 (PyPSA는 solver가 없으면 에러 발생)
            try:
                # 지정된 솔버로 먼저 시도
                if solver_name not in ['cplex', 'gurobi', 'cbc', 'glpk']:
                    logger.warning(f"지정된 솔버 '{solver_name}'는 지원되지 않는 형식입니다.")
                    solver_name = 'cplex'  # 기본 솔버로 변경
                
                # 가장 많이 사용되는 솔버 순서로 시도
                solvers_to_try = ['cplex', 'gurobi', 'cbc', 'glpk']
                
                if solver_name != solvers_to_try[0]:
                    # 지정된 솔버를 가장 앞으로 이동
                    solvers_to_try.remove(solver_name)
                    solvers_to_try.insert(0, solver_name)
                
                solver_success = False
                
                for solver in solvers_to_try:
                    try:
                        # 간단한 테스트 - 빈 네트워크로 최적화 시도
                        test_network = pypsa.Network()
                        test_network.add("Bus", "test")
                        test_network.add("Load", "test_load", bus="test", p_set=1)
                        test_network.add("Generator", "test_gen", bus="test", p_nom=1, marginal_cost=1)
                        test_network.optimize(solver_name=solver)
                        
                        # 성공하면 선택한 솔버 사용
                        solver_name = solver
                        solver_success = True
                        logger.info(f"솔버 '{solver_name}'를 사용합니다.")
                        break
                    except Exception:
                        logger.debug(f"솔버 '{solver}'를 사용할 수 없습니다.")
                        continue
                
                if not solver_success:
                    logger.error("사용 가능한 솔버가 없습니다.")
                    return False
                
            except Exception as e:
                logger.warning(f"솔버 확인 중 오류 발생: {str(e)}. 기본 솔버를 사용합니다.")
                solver_name = 'cbc'  # 대부분의 시스템에서 사용 가능한 오픈소스 솔버
            
            # 최적화 옵션 설정
            solver_options = self._prepare_solver_options(solver_name)
            
            # 최적화 시작
            logger.info(f"네트워크 최적화 시작 (솔버: {solver_name})")
            logger.debug(f"사용 코어 수: {self.num_cores}")
            logger.debug(f"솔버 옵션: {solver_options}")
            
            start_time = time.time()
            
            # 최적화 실행
            status = network.optimize(
                solver_name=solver_name,
                solver_options=solver_options,
                extra_functionality=self._add_extra_constraints
            )
            
            optimization_time = time.time() - start_time
            
            # 결과 확인
            if hasattr(network, 'objective') and network.objective is not None:
                logger.info(f"최적화 완료 (소요 시간: {optimization_time:.2f}초)")
                logger.info(f"목적함수 값: {network.objective:.2f}")
                return True
            else:
                logger.error("최적화 실패: 목적함수 값이 없습니다.")
                return False
            
        except Exception as e:
            logger.error(f"최적화 중 오류 발생: {str(e)}")
            import traceback
            logger.debug(traceback.format_exc())
            return False
    
    def _prepare_solver_options(self, solver_name=None):
        """솔버 옵션 준비
        
        Args:
            solver_name (str, optional): 사용할 솔버 이름
            
        Returns:
            dict: 솔버 옵션 딕셔너리
        """
        # 기본 솔버 옵션
        solver_options = {
            'threads': self.num_cores
        }
        
        # 솔버별 특화된 옵션 추가
        if solver_name == 'cplex':
            solver_options.update({
                'lpmethod': 4,  # Barrier method
                'barrier.algorithm': 3,
                'parallel': 1,
                'solutiontype': 2
            })
        elif solver_name == 'gurobi':
            solver_options.update({
                'Method': 2,  # Barrier method
                'Crossover': 0,
                'BarConvTol': 1e-6,
                'FeasibilityTol': 1e-6,
                'NumericFocus': 3
            })
        elif solver_name == 'cbc':
            # CBC는 고급 옵션이 제한적
            pass
        
        # 사용자 지정 옵션 적용 (기본 옵션 덮어쓰기)
        if self.solver_options:
            solver_options.update(self.solver_options)
        
        return solver_options
    
    def _add_extra_constraints(self, network, snapshots):
        """최적화 모델에 추가 제약조건 추가
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            snapshots: 시간 스냅샷
        """
        # 추가 제약조건이 필요한 경우 여기에 구현
        pass
    
    def check_temporal_constraints(self, network):
        """시간별 제약조건 충돌 검사
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            
        Returns:
            bool: 제약조건 충돌 없음 여부
        """
        try:
            # 간소화된 옵션으로 최적화 시도
            simple_options = {
                'threads': self.num_cores,
                'lpmethod': 4,
                'barrier.algorithm': 3,
                'solutiontype': 2
            }
            
            logger.info("시간별 제약조건 충돌 검사 중...")
            network.optimize(solver_name=self.solver_name, solver_options=simple_options)
            
            logger.info("제약조건 충돌 없음")
            return True
            
        except Exception as e:
            logger.warning(f"제약조건 충돌 발생: {str(e)}")
            return False
    
    def analyze_infeasibility(self, network):
        """최적화 실패 원인 분석
        
        Args:
            network (pypsa.Network): PyPSA 네트워크 객체
            
        Returns:
            dict: 실패 원인 정보
        """
        # CPLEX 솔버를 사용하는 경우
        if self.solver_name == 'cplex':
            try:
                # 특수 옵션 설정하여 다시 최적화
                infeas_options = {
                    'threads': self.num_cores,
                    'lpmethod': 4,
                    'solutiontype': 2,
                    'populatelim': 10,
                    'numericalemphasis': 1,
                    'iisfind': 1  # 실행 불가능한 제약조건 하위 집합 식별
                }
                
                logger.info("최적화 실패 원인 분석 중...")
                network.optimize(solver_name='cplex', solver_options=infeas_options)
                
                # 여기서 문제 식별이 가능하면 추가 코드 작성
                
            except Exception as e:
                logger.error(f"실패 원인 분석 중 오류: {str(e)}")
        
        # 향후 다른 솔버에 대한 분석 방법 추가 가능
        
        return {"status": "failure", "message": "실행 불가능한 모델"} 