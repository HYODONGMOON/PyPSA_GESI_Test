import pypsa
import pandas as pd
import numpy as np
import os

def create_simple_network():
    """간단한 PyPSA 네트워크 생성"""
    print("=== 간단한 PyPSA 네트워크 생성 ===")
    
    # 네트워크 생성
    network = pypsa.Network()
    
    # 시간 설정 (24시간)
    snapshots = pd.date_range('2024-01-01', periods=24, freq='h')
    network.set_snapshots(snapshots)
    
    # Carriers 정의
    carriers = ['AC', 'electricity', 'gas', 'coal', 'nuclear', 'solar', 'wind', 'hydrogen']
    for carrier in carriers:
        network.add("Carrier", carrier, co2_emissions=0.4 if carrier in ['gas', 'coal'] else 0)
    
    # 버스 추가 (주요 지역만)
    regions = ['SEL', 'BSN', 'ICN', 'DGU', 'GWJ']
    for region in regions:
        network.add("Bus", f"{region}_EL", v_nom=345, carrier='AC')
    
    # 발전기 추가
    for region in regions:
        # 원자력
        if region in ['BSN', 'GWJ']:
            network.add("Generator", f"{region}_Nuclear",
                       bus=f"{region}_EL",
                       p_nom=1000,
                       marginal_cost=10,
                       carrier='nuclear')
        
        # LNG
        network.add("Generator", f"{region}_LNG",
                   bus=f"{region}_EL",
                   p_nom=800,
                   marginal_cost=60,
                   carrier='gas')
        
        # 태양광
        network.add("Generator", f"{region}_PV",
                   bus=f"{region}_EL",
                   p_nom=500,
                   marginal_cost=0,
                   carrier='solar',
                   p_max_pu=0.3)  # 간단한 고정값
        
        # 풍력
        if region in ['GWJ', 'BSN']:
            network.add("Generator", f"{region}_WT",
                       bus=f"{region}_EL",
                       p_nom=300,
                       marginal_cost=0,
                       carrier='wind',
                       p_max_pu=0.4)  # 간단한 고정값
    
    # 부하 추가
    load_values = {'SEL': 2000, 'BSN': 1500, 'ICN': 800, 'DGU': 1000, 'GWJ': 600}
    for region, load in load_values.items():
        network.add("Load", f"{region}_Load",
                   bus=f"{region}_EL",
                   p_set=load)
    
    # 송전선 추가 (주요 연결만)
    lines = [
        ('SEL_ICN', 'SEL_EL', 'ICN_EL'),
        ('SEL_BSN', 'SEL_EL', 'BSN_EL'),
        ('BSN_DGU', 'BSN_EL', 'DGU_EL'),
        ('DGU_GWJ', 'DGU_EL', 'GWJ_EL')
    ]
    
    for line_name, bus0, bus1 in lines:
        network.add("Line", line_name,
                   bus0=bus0,
                   bus1=bus1,
                   s_nom=2000,
                   length=100,
                   x=0.1,
                   r=0.01)
    
    # 저장장치 추가 (주요 지역만)
    for region in ['SEL', 'BSN']:
        network.add("Store", f"{region}_ESS",
                   bus=f"{region}_EL",
                   e_nom=1000,
                   carrier='electricity',
                   e_cyclic=True,
                   efficiency_store=0.9,
                   efficiency_dispatch=0.9)
    
    print(f"네트워크 구성 완료:")
    print(f"- 버스: {len(network.buses)}개")
    print(f"- 발전기: {len(network.generators)}개")
    print(f"- 부하: {len(network.loads)}개")
    print(f"- 송전선: {len(network.lines)}개")
    print(f"- 저장장치: {len(network.stores)}개")
    
    return network

def run_optimization(network):
    """최적화 실행"""
    print("\n=== 최적화 시작 ===")
    
    try:
        # 네트워크 일관성 확인
        print("네트워크 일관성 확인 중...")
        
        # 최적화 실행
        status = network.optimize(solver_name='highs')  # 무료 솔버 사용
        
        # 상태 확인 - 튜플 형태도 처리
        is_optimal = False
        if isinstance(status, tuple):
            is_optimal = status[1] == "optimal"
        else:
            is_optimal = status == "optimal"
        
        if is_optimal:
            print("✓ 최적화 성공!")
            
            # 결과 출력
            print(f"\n=== 최적화 결과 ===")
            print(f"총 비용: {network.objective:.2f}")
            
            # 발전기별 총 출력
            print(f"\n발전기별 총 출력 (MWh):")
            gen_output = network.generators_t.p.sum()
            for gen_name in gen_output.index:
                print(f"  {gen_name}: {gen_output[gen_name]:.2f}")
            
            # 송전선 최대 조류
            if len(network.lines) > 0:
                print(f"\n송전선 최대 조류 (MW):")
                line_flow = network.lines_t.p0.abs().max()
                for line_name in line_flow.index:
                    print(f"  {line_name}: {line_flow[line_name]:.2f}")
            
            # 저장장치 상태
            if len(network.stores) > 0:
                print(f"\n저장장치 최종 충전량 (MWh):")
                store_energy = network.stores_t.e.iloc[-1]
                for store_name in store_energy.index:
                    print(f"  {store_name}: {store_energy[store_name]:.2f}")
            
            # 결과 저장
            save_simple_results(network)
            
            return True
            
        else:
            print(f"✗ 최적화 실패: {status}")
            return False
            
    except Exception as e:
        print(f"최적화 중 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def save_simple_results(network):
    """간단한 결과 저장"""
    try:
        print("\n결과 저장 중...")
        
        # 결과 폴더 생성
        if not os.path.exists('simple_results'):
            os.makedirs('simple_results')
        
        # 발전기 출력 저장
        network.generators_t.p.to_csv('simple_results/generator_output.csv')
        
        # 부하 저장
        network.loads_t.p.to_csv('simple_results/load.csv')
        
        # 송전선 조류 저장
        if len(network.lines) > 0:
            network.lines_t.p0.to_csv('simple_results/line_flow.csv')
        
        # 저장장치 상태 저장
        if len(network.stores) > 0:
            network.stores_t.e.to_csv('simple_results/storage_energy.csv')
            network.stores_t.p.to_csv('simple_results/storage_power.csv')
        
        # 요약 정보 저장
        summary = {
            'total_cost': network.objective,
            'total_generation': network.generators_t.p.sum().sum(),
            'total_load': network.loads_t.p.sum().sum(),
            'optimization_status': 'optimal'
        }
        
        summary_df = pd.DataFrame([summary])
        summary_df.to_csv('simple_results/summary.csv', index=False)
        
        print("결과가 'simple_results' 폴더에 저장되었습니다.")
        
    except Exception as e:
        print(f"결과 저장 중 오류: {str(e)}")

def main():
    """메인 함수"""
    print("=== 간단한 PyPSA 분석 시작 ===")
    
    # 네트워크 생성
    network = create_simple_network()
    
    # 최적화 실행
    success = run_optimization(network)
    
    if success:
        print("\n✓ 분석이 성공적으로 완료되었습니다!")
    else:
        print("\n✗ 분석 중 오류가 발생했습니다.")

if __name__ == "__main__":
    main() 