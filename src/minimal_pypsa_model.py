import pypsa
import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
import sys
import traceback

def create_simple_model():
    """
    매우 간단한 PyPSA 모델 생성 및 실행
    """
    print("간단한 PyPSA 모델 생성 중...")
    print(f"Python 버전: {sys.version}")
    print(f"PyPSA 버전: {pypsa.__version__}")
    
    # 새 네트워크 생성
    try:
        network = pypsa.Network()
        
        # 단일 스냅샷 설정
        network.set_snapshots(pd.date_range("2023-01-01 00:00", "2023-01-01 01:00", freq="1h"))
        print("스냅샷 설정 완료")
        
        # 기본 버스 추가
        for region in ["SEL", "BSN", "ICN", "GGD"]:
            network.add("Bus", f"{region}_EL", v_nom=380, carrier="AC")
        
        # 외부 그리드 버스 추가
        network.add("Bus", "External_Grid", v_nom=380, carrier="AC")
        print("버스 추가 완료")
        
        # 주요 지역 부하 추가
        loads = {
            "SEL_EL": 8000,  # MW
            "BSN_EL": 6000,  # MW
            "ICN_EL": 6000,  # MW
            "GGD_EL": 10000  # MW
        }
        
        for bus, p in loads.items():
            network.add("Load", f"{bus.split('_')[0]}_load", bus=bus, p_set=p)
        print("부하 추가 완료")
        
        # 각 지역 발전기 추가
        generators = {
            "SEL_EL": {"p_nom": 5000, "marginal_cost": 50},
            "BSN_EL": {"p_nom": 8000, "marginal_cost": 40},
            "ICN_EL": {"p_nom": 7000, "marginal_cost": 45},
            "GGD_EL": {"p_nom": 9000, "marginal_cost": 55}
        }
        
        for bus, params in generators.items():
            network.add("Generator", 
                       f"{bus.split('_')[0]}_gen", 
                       bus=bus, 
                       p_nom=params["p_nom"],
                       p_nom_extendable=True,
                       p_nom_min=0,
                       p_nom_max=params["p_nom"] * 2,
                       marginal_cost=params["marginal_cost"])
        
        # 외부 그리드 발전기 추가 (무제한 발전)
        network.add("Generator",
                   "external_gen",
                   bus="External_Grid",
                   p_nom=100000,
                   p_nom_extendable=True,
                   marginal_cost=100)  # 비싼 발전 비용
        print("발전기 추가 완료")
        
        # 지역간 선로 연결
        lines = [
            {"bus0": "SEL_EL", "bus1": "GGD_EL", "s_nom": 5000},
            {"bus0": "BSN_EL", "bus1": "GGD_EL", "s_nom": 3000},
            {"bus0": "ICN_EL", "bus1": "SEL_EL", "s_nom": 4000},
            {"bus0": "ICN_EL", "bus1": "GGD_EL", "s_nom": 3000}
        ]
        
        for i, line_data in enumerate(lines):
            network.add("Line",
                       f"line_{i+1}",
                       bus0=line_data["bus0"],
                       bus1=line_data["bus1"],
                       s_nom=line_data["s_nom"],
                       s_nom_extendable=True,
                       s_nom_max=line_data["s_nom"] * 3,
                       x=0.2,  # 리액턴스
                       r=0.05)  # 저항
        print("선로 추가 완료")
        
        # 수입/수출 링크 추가
        for region_bus in ["SEL_EL", "BSN_EL", "ICN_EL", "GGD_EL"]:
            region = region_bus.split('_')[0]
            
            # 수입 링크
            network.add("Link",
                       f"import_{region}",
                       bus0="External_Grid",
                       bus1=region_bus,
                       p_nom=5000,
                       p_nom_extendable=True,
                       p_nom_max=10000,
                       efficiency=1.0,
                       marginal_cost=90)
            
            # 수출 링크
            network.add("Link",
                       f"export_{region}",
                       bus0=region_bus,
                       bus1="External_Grid",
                       p_nom=3000,
                       p_nom_extendable=True,
                       p_nom_max=8000,
                       efficiency=0.95,
                       marginal_cost=-10)  # 수출은 음의 비용
        print("링크 추가 완료")
        
        # 저장장치 추가 (간단한 배터리)
        for region_bus in ["SEL_EL", "BSN_EL"]:
            region = region_bus.split('_')[0]
            network.add("Store",
                       f"{region}_battery",
                       bus=region_bus,
                       e_nom=1000,  # MWh
                       e_nom_extendable=True,
                       e_cyclic=True,
                       standing_loss=0.01)  # 1%/h 자연 손실
        print("저장장치 추가 완료")
        
        print("네트워크 모델 생성 완료")
        print(f"버스: {len(network.buses)}개")
        print(f"발전기: {len(network.generators)}개")
        print(f"부하: {len(network.loads)}개")
        print(f"선로: {len(network.lines)}개")
        print(f"링크: {len(network.links)}개")
        print(f"저장장치: {len(network.stores)}개")
        
        # 최적화 실행
        print("\n최적화 시작...")
        
        try:
            # 네트워크 저장 (시각화 목적)
            network.export_to_netcdf("simple_network_before.nc")
            print("네트워크 데이터가 simple_network_before.nc에 저장되었습니다.")
        except Exception as e:
            print(f"네트워크 저장 오류: {e}")
        
        # 사용 가능한 솔버 확인
        try:
            available_solvers = pypsa.available_solvers()
            print(f"사용 가능한 솔버: {available_solvers}")
        except:
            print("사용 가능한 솔버를 확인할 수 없습니다.")
        
        try:
            # 최적화 옵션 설정
            solver_options = {
                "threads": 4,
                "method": 1,  # dual simplex
                "crossover": 0,
                "BarConvTol": 1.e-5,
                "FeasibilityTol": 1.e-6,
                "AggFill": 0,
                "PreDual": 0,
                "PRELINEAR": 0
            }
            
            # 최적화 실행
            print("GLPK 솔버로 최적화 시도 중...")
            network.lopf(solver_name="glpk")
            
            print("최적화 성공!")
            
            # 결과 요약
            results = {
                "objective": network.objective,
                "generators": network.generators_t.p.sum(),
                "loads": network.loads_t.p.sum(),
                "lines": {line: network.lines_t.p0.loc[:, line].sum() for line in network.lines.index}
            }
            
            print("\n최적화 결과 요약:")
            print(f"목적 함수 값: {network.objective:.2f}")
            print("\n발전기 출력 (MWh):")
            for gen, value in results["generators"].items():
                print(f"  - {gen}: {value:.2f}")
            
            print("\n선로 흐름 (MWh):")
            for line, value in results["lines"].items():
                bus0 = network.lines.loc[line, "bus0"]
                bus1 = network.lines.loc[line, "bus1"]
                print(f"  - {line} ({bus0} → {bus1}): {value:.2f}")
            
            # 결과 네트워크 저장
            network.export_to_netcdf("simple_network_optimized.nc")
            
            # 그래프 시각화 (선택 사항)
            try:
                plt.figure(figsize=(12, 8))
                
                # 발전량 시각화
                plt.subplot(2, 1, 1)
                network.generators_t.p.plot(title="발전기 출력")
                plt.ylabel("출력 (MW)")
                plt.grid(True)
                
                # 선로 흐름 시각화
                plt.subplot(2, 1, 2)
                network.lines_t.p0.plot(title="선로 흐름")
                plt.ylabel("흐름 (MW)")
                plt.grid(True)
                
                plt.tight_layout()
                plt.savefig("optimization_results.png")
                print("\n결과가 'optimization_results.png'에 저장되었습니다.")
                
            except Exception as e:
                print(f"그래프 생성 중 오류 발생: {e}")
                traceback.print_exc()
            
            return True
            
        except Exception as e:
            print(f"최적화 중 오류 발생: {e}")
            traceback.print_exc()
            
            # CBC 솔버 시도
            try:
                print("\nCBC 솔버로 다시 시도 중...")
                network.lopf(solver_name="cbc")
                print("CBC 최적화 성공!")
                return True
            except Exception as e2:
                print(f"CBC 솔버 오류: {e2}")
                traceback.print_exc()
            
            return False
    
    except Exception as e:
        print(f"모델 생성 중 오류 발생: {e}")
        traceback.print_exc()
        return False

if __name__ == "__main__":
    result = create_simple_model()
    print(f"\n최종 결과: {'성공' if result else '실패'}") 