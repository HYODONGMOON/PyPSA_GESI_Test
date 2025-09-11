import pypsa
import pandas as pd
import sys
import traceback
import os

# 로그 파일 설정
log_file = 'pypsa_cplex_log.txt'

# 로그 파일 생성 또는 초기화
with open(log_file, 'w') as f:
    f.write('PyPSA CPLEX 실행 로그\n')
    f.write('=' * 50 + '\n\n')

def log(message):
    """로그 메시지를 파일과 콘솔에 출력합니다."""
    print(message)
    with open(log_file, 'a') as f:
        f.write(message + '\n')

def create_simple_model():
    """
    매우 단순한 PyPSA 모델을 생성하고 CPLEX 솔버로 최적화합니다.
    """
    try:
        log("매우 단순한 PyPSA 모델 생성 중...")
        log(f"Python 버전: {sys.version}")
        log(f"PyPSA 버전: {pypsa.__version__}")
        
        # 새 네트워크 생성
        network = pypsa.Network()
        log("네트워크 객체가 생성되었습니다.")
        
        # 단일 스냅샷 설정
        network.set_snapshots(pd.date_range("2023-01-01 00:00", "2023-01-01 01:00", freq="1h"))
        log("스냅샷이 설정되었습니다.")
        
        # 두 개의 버스 추가 (carrier 명시)
        network.add("Bus", "Bus1", v_nom=380, carrier="AC")
        network.add("Bus", "Bus2", v_nom=380, carrier="AC")
        log("두 개의 버스가 추가되었습니다.")
        
        # 부하 추가
        network.add("Load", "Load1", bus="Bus2", p_set=100)
        log("부하가 추가되었습니다.")
        
        # 발전기 추가
        network.add("Generator", "Gen1", bus="Bus1", p_nom=150, marginal_cost=50, carrier="AC")
        log("발전기가 추가되었습니다.")
        
        # 선로 추가 (carrier 명시)
        network.add("Line", "Line1", 
                   bus0="Bus1", 
                   bus1="Bus2", 
                   x=0.1, 
                   r=0.01, 
                   s_nom=200,
                   carrier="AC")
        log("선로가 추가되었습니다.")
        
        log("\n네트워크 구성 요약:")
        log(f"버스: {len(network.buses)}개")
        log(f"발전기: {len(network.generators)}개")
        log(f"부하: {len(network.loads)}개")
        log(f"선로: {len(network.lines)}개")
        
        # 최적화 실행
        log("\n최적화 시작...")
        
        try:
            # 솔버 목록 출력 (가능하다면)
            try:
                from linopy.solvers import available_solvers
                log(f"사용 가능한 솔버: {available_solvers()}")
            except Exception as e:
                log(f"솔버 목록을 가져올 수 없습니다: {e}")
            
            # 최적화 옵션 설정 (단순화)
            solver_options = {
                "threads": 4,
                "lpmethod": 4,  # 배리어 메서드
                "barrier.algorithm": 3,  # 기본 배리어 알고리즘
                # solutiontype 옵션 제거 (문제 발생)
            }
            
            log("CPLEX 솔버로 최적화 시도 중...")
            
            # 최적화 시도
            network.optimize(solver_name="cplex", 
                           solver_options=solver_options)
            
            log("최적화 성공!")
            
            # 결과 출력
            log("\n최적화 결과:")
            if hasattr(network, 'objective'):
                log(f"목적 함수 값: {network.objective}")
            else:
                log("목적 함수 값을 가져올 수 없습니다.")
            
            log("\n발전기 출력:")
            try:
                for gen in network.generators.index:
                    try:
                        output = network.generators_t.p.loc[:, gen].iloc[0]
                        log(f"  - {gen}: {output}")
                    except Exception as e:
                        log(f"  - {gen}: 출력을 가져올 수 없습니다. 오류: {e}")
            except Exception as e:
                log(f"발전기 출력 정보 접근 오류: {e}")
            
            log("\n선로 흐름:")
            try:
                for line in network.lines.index:
                    try:
                        flow = network.lines_t.p0.loc[:, line].iloc[0]
                        log(f"  - {line}: {flow}")
                    except Exception as e:
                        log(f"  - {line}: 흐름을 가져올 수 없습니다. 오류: {e}")
            except Exception as e:
                log(f"선로 흐름 정보 접근 오류: {e}")
            
            # 다른 솔버 시도
            return True
            
        except Exception as e:
            log(f"CPLEX 최적화 오류: {e}")
            traceback.print_exc(file=open(log_file, 'a'))
            
            # CBC 솔버로 재시도
            try:
                log("\nCBC 솔버로 다시 시도 중...")
                network.optimize(solver_name="cbc")
                log("CBC 최적화 성공!")
                return True
            except Exception as e2:
                log(f"CBC 솔버 오류: {e2}")
                traceback.print_exc(file=open(log_file, 'a'))
                
                # GLPK 솔버로 재시도
                try:
                    log("\nGLPK 솔버로 마지막 시도...")
                    network.optimize(solver_name="glpk")
                    log("GLPK 최적화 성공!")
                    return True
                except Exception as e3:
                    log(f"GLPK 솔버 오류: {e3}")
                    traceback.print_exc(file=open(log_file, 'a'))
                    return False
            
        
    except Exception as e:
        log(f"모델 생성 오류: {e}")
        traceback.print_exc(file=open(log_file, 'a'))
        return False

if __name__ == "__main__":
    result = create_simple_model()
    status = "성공" if result else "실패"
    log(f"\n최종 실행 결과: {status}")
    log(f"로그 파일이 {os.path.abspath(log_file)}에 저장되었습니다.") 