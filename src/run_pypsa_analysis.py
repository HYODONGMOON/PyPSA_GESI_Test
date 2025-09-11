import pypsa
import pandas as pd
import numpy as np
import os

def run_pypsa_analysis():
    print("=== PyPSA 분석 시작 ===")
    
    try:
        # 네트워크 생성
        network = pypsa.Network()
        
        # integrated_input_data.xlsx 파일에서 데이터 로드
        if os.path.exists('integrated_input_data.xlsx'):
            print("integrated_input_data.xlsx 파일에서 데이터를 로드합니다...")
            
            # 각 시트 읽기
            buses_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='buses')
            generators_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
            loads_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='loads')
            stores_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='stores')
            links_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='links')
            lines_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='lines')
            constraints_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='constraints')
            
            print(f"로드된 데이터:")
            print(f"- 버스: {len(buses_df)}개")
            print(f"- 발전기: {len(generators_df)}개")
            print(f"- 부하: {len(loads_df)}개")
            print(f"- 저장장치: {len(stores_df)}개")
            print(f"- 링크: {len(links_df)}개")
            print(f"- 송전선: {len(lines_df)}개")
            print(f"- 제약조건: {len(constraints_df)}개")
            
            # 네트워크에 컴포넌트 추가
            # 버스 추가
            for _, bus in buses_df.iterrows():
                network.add("Bus", bus['name'], 
                           v_nom=bus.get('v_nom', 380),
                           x=bus.get('x', 0),
                           y=bus.get('y', 0))
            
            # 발전기 추가
            for _, gen in generators_df.iterrows():
                network.add("Generator", gen['name'],
                           bus=gen['bus'],
                           p_nom=gen.get('p_nom', 100),
                           marginal_cost=gen.get('marginal_cost', 50),
                           carrier=gen.get('carrier', 'gas'))
            
            # 부하 추가
            for _, load in loads_df.iterrows():
                network.add("Load", load['name'],
                           bus=load['bus'],
                           p_set=load.get('p_set', 100))
            
            # 저장장치 추가
            for _, store in stores_df.iterrows():
                network.add("Store", store['name'],
                           bus=store['bus'],
                           e_nom=store.get('e_nom', 1000),
                           carrier=store.get('carrier', 'battery'))
            
            # 링크 추가
            for _, link in links_df.iterrows():
                network.add("Link", link['name'],
                           bus0=link['bus0'],
                           bus1=link['bus1'],
                           p_nom=link.get('p_nom', 100),
                           efficiency=link.get('efficiency', 0.9))
            
            # 송전선 추가
            for _, line in lines_df.iterrows():
                network.add("Line", line['name'],
                           bus0=line['bus0'],
                           bus1=line['bus1'],
                           s_nom=line.get('s_nom', 1000),
                           length=line.get('length', 100),
                           x=line.get('x', 0.1),
                           r=line.get('r', 0.01))
            
            print("\n네트워크 구성 완료:")
            print(f"- 버스: {len(network.buses)}개")
            print(f"- 발전기: {len(network.generators)}개")
            print(f"- 부하: {len(network.loads)}개")
            print(f"- 저장장치: {len(network.stores)}개")
            print(f"- 링크: {len(network.links)}개")
            print(f"- 송전선: {len(network.lines)}개")
            
            # 최적화 실행
            print("\n=== 최적화 시작 ===")
            status = network.optimize()
            
            if status == "optimal":
                print("✓ 최적화 성공!")
                
                # 결과 출력
                print("\n=== 최적화 결과 ===")
                print(f"총 비용: {network.objective:.2f}")
                
                # 발전기별 출력
                print("\n발전기 출력:")
                for gen_name in network.generators.index:
                    p_opt = network.generators_t.p[gen_name].iloc[0] if len(network.generators_t.p) > 0 else 0
                    print(f"  {gen_name}: {p_opt:.2f} MW")
                
                # 송전선 조류
                if len(network.lines) > 0:
                    print("\n송전선 조류:")
                    for line_name in network.lines.index:
                        p_flow = network.lines_t.p0[line_name].iloc[0] if len(network.lines_t.p0) > 0 else 0
                        print(f"  {line_name}: {p_flow:.2f} MW")
                
                # 결과를 CSV로 저장
                network.export_to_csv_folder("pypsa_results")
                print("\n결과가 'pypsa_results' 폴더에 저장되었습니다.")
                
            else:
                print(f"✗ 최적화 실패: {status}")
                
        else:
            print("integrated_input_data.xlsx 파일을 찾을 수 없습니다.")
            
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    run_pypsa_analysis() 