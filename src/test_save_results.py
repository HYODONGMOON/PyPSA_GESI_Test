import os
import json
import numpy as np
import pandas as pd
from datetime import datetime

# 테스트용 간단한 네트워크 생성
class MockNetwork:
    def __init__(self):
        self.objective = 1000.0
        
        # 시간 인덱스 생성
        snapshots = pd.date_range('2023-01-01', periods=24, freq='H')
        
        # 버스 설정
        self.buses = pd.DataFrame({
            'v_nom': [380.0, 380.0, 380.0],
            'carrier': ['AC', 'AC', 'AC']
        }, index=['SEL_AC', 'BSN_AC', 'GWD_AC'])
        
        # 발전기 설정
        self.generators = pd.DataFrame({
            'bus': ['SEL_AC', 'BSN_AC', 'GWD_AC'],
            'p_nom': [1000, 500, 800],
            'carrier': ['coal', 'nuclear', 'wind']
        }, index=['SEL_G1', 'BSN_G1', 'GWD_G1'])
        
        # 부하 설정
        self.loads = pd.DataFrame({
            'bus': ['SEL_AC', 'BSN_AC', 'GWD_AC'],
            'p_set': [700, 300, 400]
        }, index=['SEL_L', 'BSN_L', 'GWD_L'])
        
        # 선로 설정
        self.lines = pd.DataFrame({
            'bus0': ['SEL_AC', 'BSN_AC'],
            'bus1': ['BSN_AC', 'GWD_AC'],
            's_nom': [1000, 800]
        }, index=['Line1', 'Line2'])
        
        # 저장장치 설정
        self.stores = pd.DataFrame({
            'bus': ['SEL_AC'],
            'e_nom': [500]
        }, index=['SEL_S1'])
        
        # 시계열 데이터 설정
        self.generators_t = type('', (), {})()
        self.generators_t.p = pd.DataFrame(
            np.random.rand(24, 3) * np.array([[800, 400, 600]]),
            index=snapshots,
            columns=['SEL_G1', 'BSN_G1', 'GWD_G1']
        )
        
        self.loads_t = type('', (), {})()
        self.loads_t.p = pd.DataFrame(
            np.random.rand(24, 3) * np.array([[600, 250, 350]]),
            index=snapshots,
            columns=['SEL_L', 'BSN_L', 'GWD_L']
        )
        
        self.lines_t = type('', (), {})()
        self.lines_t.p0 = pd.DataFrame(
            np.random.rand(24, 2) * np.array([[500, 400]]),
            index=snapshots,
            columns=['Line1', 'Line2']
        )
        
        self.stores_t = type('', (), {})()
        self.stores_t.e = pd.DataFrame(
            np.random.rand(24, 1) * np.array([[300]]),
            index=snapshots,
            columns=['SEL_S1']
        )
        
        # 메타데이터 설정
        self.meta = {
            'solver': {
                'name': 'mockup_solver',
                'time': 10.5
            }
        }
    
    def export_to_netcdf(self, filename):
        """네트워크를 넷CDF 파일로 저장하는 척합니다."""
        print(f"네트워크를 '{filename}'에 넷CDF 형식으로 저장했습니다 (목업).")
        # 파일 생성
        with open(filename, 'w') as f:
            f.write("This is a mockup netCDF file.")

def save_results(network):
    """네트워크 최적화 결과를 저장합니다."""
    if network is None:
        print("네트워크가 없어 결과를 저장할 수 없습니다.")
        return
    
    try:
        # 결과 저장 디렉토리 생성
        results_dir = "results"
        os.makedirs(results_dir, exist_ok=True)
        
        # 현재 시간을 기반으로 고유한 파일 이름 생성
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # 분석 일시를 이름으로 하는 폴더 생성
        timestamp_dir = os.path.join(results_dir, timestamp)
        os.makedirs(timestamp_dir, exist_ok=True)
        print(f"결과 폴더 '{timestamp_dir}'을 생성했습니다.")
        
        # 파일 이름 기본 경로 설정
        filename_base = os.path.join(timestamp_dir, f"optimization_result_{timestamp}")
        
        # 네트워크 객체 저장
        network_file = f"{filename_base}.nc"
        network.export_to_netcdf(network_file)
        print(f"네트워크 결과가 '{network_file}'에 저장되었습니다.")
        
        # 최적화 통계 저장
        if hasattr(network, 'objective'):
            stats = {
                "objective_value": float(network.objective) if hasattr(network, 'objective') else None,
                "total_generation": float(network.generators_t.p.sum().sum()) if hasattr(network, 'generators_t') and hasattr(network.generators_t, 'p') else None,
                "total_load": float(network.loads_t.p.sum().sum()) if hasattr(network, 'loads_t') and hasattr(network.loads_t, 'p') else None,
                "timestamp": timestamp,
                "solver": network.meta.get('solver', {}).get('name', 'unknown') if hasattr(network, 'meta') else 'unknown',
                "solving_time": network.meta.get('solver', {}).get('time', 0) if hasattr(network, 'meta') else 0
            }
            
            # JSON으로 통계 저장
            stats_file = f"{filename_base}_stats.json"
            with open(stats_file, 'w') as f:
                json.dump(stats, f, indent=4)
            print(f"최적화 통계가 '{stats_file}'에 저장되었습니다.")
            
            # 발전기 출력 저장
            if hasattr(network, 'generators_t') and hasattr(network.generators_t, 'p'):
                gen_output_file = f"{filename_base}_generator_output.csv"
                network.generators_t.p.to_csv(gen_output_file)
                print(f"발전기 출력이 '{gen_output_file}'에 저장되었습니다.")
            
            # 부하 데이터 저장
            if hasattr(network, 'loads_t') and hasattr(network.loads_t, 'p'):
                load_file = f"{filename_base}_load.csv"
                network.loads_t.p.to_csv(load_file)
                print(f"부하 데이터가 '{load_file}'에 저장되었습니다.")
            
            # 라인 사용률 저장
            if hasattr(network, 'lines_t') and hasattr(network.lines_t, 'p0'):
                lines_file = f"{filename_base}_line_usage.csv"
                network.lines_t.p0.to_csv(lines_file)
                print(f"라인 사용률이 '{lines_file}'에 저장되었습니다.")
            
            # 저장소 사용률 저장
            if hasattr(network, 'stores_t') and hasattr(network.stores_t, 'e'):
                stores_file = f"{filename_base}_storage.csv"
                network.stores_t.e.to_csv(stores_file)
                print(f"저장소 에너지 레벨이 '{stores_file}'에 저장되었습니다.")
        
        # 분석을 위해 최신 타임스탬프 저장
        latest_timestamp_file = os.path.join(results_dir, "latest_timestamp.txt")
        with open(latest_timestamp_file, 'w') as f:
            f.write(timestamp)
        print(f"최신 타임스탬프가 '{latest_timestamp_file}'에 저장되었습니다.")
        
        print("\n지역별 분석 및 시각화 이부분은 실제 PyPSA_GUI.py에서 진행됩니다.")
        print(f"(테스트 스크립트에서는 생략합니다)")
        
        return True
    except Exception as e:
        print(f"결과 저장 중 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    # 테스트용 네트워크 생성
    network = MockNetwork()
    
    print("테스트 네트워크 생성 완료. save_results 함수 실행...")
    
    # save_results 함수 호출
    result = save_results(network)
    
    if result:
        print("테스트 성공: 모든 결과가 성공적으로 저장되었습니다.")
    else:
        print("테스트 실패: 결과 저장 중 오류가 발생했습니다.") 