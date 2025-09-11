import pypsa
import pandas as pd
import sys
import traceback
import os

# 로그 파일 설정
log_file = 'two_region_model_log.txt'

# 로그 파일 생성 또는 초기화
with open(log_file, 'w') as f:
    f.write('두 지역 간 연결 모델 실행 로그\n')
    f.write('=' * 50 + '\n\n')

def log(message):
    """로그 메시지를 파일과 콘솔에 출력합니다."""
    print(message)
    with open(log_file, 'a') as f:
        f.write(message + '\n')

def create_two_region_model():
    """
    두 지역(서울, 부산) 간 연결을 가진 단순한 모델을 생성합니다.
    성공적인 CPLEX 설정을 사용합니다.
    """
    try:
        log("두 지역 간 연결 모델 생성 중...")
        log(f"Python 버전: {sys.version}")
        log(f"PyPSA 버전: {pypsa.__version__}")
        
        # 새 네트워크 생성
        network = pypsa.Network()
        log("네트워크 객체가 생성되었습니다.")
        
        # 단일 스냅샷 설정 (24시간 기간으로 확장)
        network.set_snapshots(pd.date_range("2023-01-01 00:00", "2023-01-01 23:00", freq="1h"))
        log(f"스냅샷이 설정되었습니다: {len(network.snapshots)}개")
        
        # 지역 정의
        regions = {
            "SEL": {"name": "서울", "load": 8000, "gen_capacity": 5000, "gen_cost": 50},
            "BSN": {"name": "부산", "load": 6000, "gen_capacity": 8000, "gen_cost": 40}
        }
        
        # 버스 추가
        for code, data in regions.items():
            network.add("Bus", f"{code}_EL", v_nom=380, carrier="AC", x=0, y=0)
        
        # 외부 그리드 추가
        network.add("Bus", "External_Grid", v_nom=380, carrier="AC", x=0, y=0)
        log("모든 버스가 추가되었습니다.")
        
        # 부하 추가 (시간에 따라 변화하는 패턴)
        load_pattern = pd.Series([0.7, 0.65, 0.6, 0.55, 0.55, 0.6, 0.7, 0.8, 0.9, 1.0, 
                                   0.95, 0.9, 0.9, 0.85, 0.85, 0.9, 0.95, 1.0, 
                                   1.0, 0.95, 0.9, 0.85, 0.8, 0.75], index=network.snapshots)
        
        for code, data in regions.items():
            load_ts = pd.DataFrame(index=network.snapshots, 
                                   data={f"{code}_load": data["load"] * load_pattern.values})
            
            network.add("Load", f"{code}_load", 
                       bus=f"{code}_EL", 
                       p_set=load_ts[f"{code}_load"])
        log("모든 부하가 추가되었습니다.")
        
        # 발전기 추가
        for code, data in regions.items():
            network.add("Generator", 
                       f"{code}_gen", 
                       bus=f"{code}_EL", 
                       p_nom=data["gen_capacity"],
                       p_nom_extendable=True,
                       p_nom_min=0,
                       p_nom_max=data["gen_capacity"] * 2,
                       marginal_cost=data["gen_cost"],
                       carrier="AC")
        
        # 외부 그리드 발전기 추가
        network.add("Generator",
                   "external_gen",
                   bus="External_Grid",
                   p_nom=20000,
                   p_nom_extendable=True,
                   marginal_cost=100,
                   carrier="AC")
        log("모든 발전기가 추가되었습니다.")
        
        # 지역 간 선로 추가
        network.add("Line", "SEL_BSN_line", 
                   bus0="SEL_EL", 
                   bus1="BSN_EL", 
                   x=0.2, 
                   r=0.05, 
                   s_nom=5000,
                   s_nom_extendable=True,
                   carrier="AC")
        log("지역 간 선로가 추가되었습니다.")
        
        # 외부 그리드 링크 추가
        for code in regions.keys():
            # 수입 링크
            network.add("Link",
                       f"import_{code}",
                       bus0="External_Grid",
                       bus1=f"{code}_EL",
                       p_nom=5000,
                       p_nom_extendable=True,
                       p_nom_max=10000,
                       efficiency=1.0,
                       marginal_cost=90,
                       carrier="AC")
            
            # 수출 링크
            network.add("Link",
                       f"export_{code}",
                       bus0=f"{code}_EL",
                       bus1="External_Grid",
                       p_nom=3000,
                       p_nom_extendable=True,
                       p_nom_max=8000,
                       efficiency=0.95,
                       marginal_cost=-10,
                       carrier="AC")
        log("외부 그리드 연결 링크가 추가되었습니다.")
        
        # 저장장치 추가
        for code in regions.keys():
            network.add("Store",
                       f"{code}_battery",
                       bus=f"{code}_EL",
                       e_nom=1000,
                       e_nom_extendable=True,
                       e_cyclic=True,
                       standing_loss=0.01,
                       carrier="battery")
        log("저장장치가 추가되었습니다.")
        
        log("\n네트워크 구성 요약:")
        log(f"버스: {len(network.buses)}개")
        log(f"발전기: {len(network.generators)}개")
        log(f"부하: {len(network.loads)}개")
        log(f"선로: {len(network.lines)}개")
        log(f"링크: {len(network.links)}개")
        log(f"저장장치: {len(network.stores)}개")
        log(f"스냅샷: {len(network.snapshots)}개")
        
        # 최적화 실행
        log("\n최적화 시작...")
        
        try:
            # 최적화 옵션 설정
            solver_options = {
                "threads": 4,
                "lpmethod": 4,  # 배리어 메서드
                "barrier.algorithm": 3  # 기본 배리어 알고리즘
            }
            
            log("CPLEX 솔버로 최적화 시도 중...")
            
            # 최적화 실행
            network.optimize(solver_name="cplex", 
                          solver_options=solver_options)
            
            log("최적화 성공!")
            
            # 결과 요약
            log("\n최적화 결과:")
            log(f"목적 함수 값: {network.objective:.2f}")
            
            # 지역별 발전량 계산
            log("\n지역별 총 발전량 (MWh):")
            for code in regions.keys():
                gen_name = f"{code}_gen"
                total_gen = network.generators_t.p[gen_name].sum()
                log(f"  - {regions[code]['name']}: {total_gen:.2f}")
            
            external_gen = network.generators_t.p["external_gen"].sum()
            log(f"  - 외부 그리드: {external_gen:.2f}")
            
            # 지역별 부하량 계산
            log("\n지역별 총 부하량 (MWh):")
            for code in regions.keys():
                load_name = f"{code}_load"
                total_load = network.loads_t.p[load_name].sum()
                log(f"  - {regions[code]['name']}: {total_load:.2f}")
            
            # 선로 흐름 분석
            log("\n지역 간 선로 흐름 분석:")
            line_name = "SEL_BSN_line"
            max_flow = network.lines_t.p0[line_name].abs().max()
            avg_flow = network.lines_t.p0[line_name].mean()
            sel_to_bsn = network.lines_t.p0[line_name][network.lines_t.p0[line_name] > 0].sum()
            bsn_to_sel = network.lines_t.p0[line_name][network.lines_t.p0[line_name] < 0].sum() * -1
            
            log(f"  - 최대 흐름: {max_flow:.2f} MW")
            log(f"  - 평균 흐름: {avg_flow:.2f} MW")
            log(f"  - 서울 → 부산: {sel_to_bsn:.2f} MWh")
            log(f"  - 부산 → 서울: {bsn_to_sel:.2f} MWh")
            
            # 저장장치 사용 분석
            log("\n저장장치 사용 분석:")
            for code in regions.keys():
                store_name = f"{code}_battery"
                if store_name in network.stores_t.p:
                    max_charge = network.stores_t.p[store_name].max()
                    max_discharge = network.stores_t.p[store_name].min() * -1
                    net_energy = network.stores_t.p[store_name].sum()
                    log(f"  - {regions[code]['name']} 배터리:")
                    log(f"    - 최대 충전: {max_charge:.2f} MW")
                    log(f"    - 최대 방전: {max_discharge:.2f} MW")
                    log(f"    - 순 에너지: {net_energy:.2f} MWh")
                else:
                    log(f"  - {regions[code]['name']} 배터리: 데이터 없음")
            
            # 결과 저장
            try:
                network.export_to_netcdf("two_region_model_results.nc")
                log("결과가 'two_region_model_results.nc'에 저장되었습니다.")
            except Exception as e:
                log(f"결과 저장 오류: {e}")
            
            return True
            
        except Exception as e:
            log(f"CPLEX 최적화 오류: {e}")
            traceback.print_exc(file=open(log_file, 'a'))
            
            # 다른 솔버 시도
            try:
                log("\nGLPK 솔버로 다시 시도...")
                network.optimize(solver_name="glpk")
                log("GLPK 최적화 성공!")
                return True
            except Exception as e2:
                log(f"GLPK 솔버 오류: {e2}")
                traceback.print_exc(file=open(log_file, 'a'))
                return False
        
    except Exception as e:
        log(f"모델 생성 오류: {e}")
        traceback.print_exc(file=open(log_file, 'a'))
        return False

if __name__ == "__main__":
    result = create_two_region_model()
    status = "성공" if result else "실패"
    log(f"\n최종 실행 결과: {status}")
    log(f"로그 파일이 {os.path.abspath(log_file)}에 저장되었습니다.") 