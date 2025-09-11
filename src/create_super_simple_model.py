import pandas as pd
import os
from datetime import datetime

def create_super_simple_model():
    """
    매우 단순한 테스트 모델을 생성합니다.
    이 모델은 2개의 지역과 외부 연결 그리드를 포함합니다.
    """
    print("매우 단순한 테스트 모델 생성 중...")
    
    # Excel 파일 경로 설정
    output_file = 'super_simple_model.xlsx'
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    try:
        # 지역 코드 (서울, 부산)
        regions = ['SEL', 'BSN']
        
        # 필요한 데이터프레임 생성
        # 버스
        buses = pd.DataFrame([
            # 서울 지역 버스
            {'name': 'SEL_EL', 'v_nom': 345, 'carrier': 'AC', 'x': 100, 'y': 100},
            # 부산 지역 버스
            {'name': 'BSN_EL', 'v_nom': 345, 'carrier': 'AC', 'x': 300, 'y': 100},
            # 외부 그리드
            {'name': 'External_Grid', 'v_nom': 345, 'carrier': 'AC', 'x': 0, 'y': 0}
        ])
        
        # 발전기
        generators = pd.DataFrame([
            # 서울 지역 발전기
            {'name': 'SEL_Gen1', 'bus': 'SEL_EL', 'p_nom': 100, 'p_nom_extendable': True, 
             'p_nom_min': 0, 'p_nom_max': 500, 'marginal_cost': 50, 'carrier': 'AC'},
            # 부산 지역 발전기
            {'name': 'BSN_Gen1', 'bus': 'BSN_EL', 'p_nom': 100, 'p_nom_extendable': True, 
             'p_nom_min': 0, 'p_nom_max': 500, 'marginal_cost': 60, 'carrier': 'AC'}
        ])
        
        # 부하
        loads = pd.DataFrame([
            # 서울 지역 부하
            {'name': 'SEL_Load', 'bus': 'SEL_EL', 'p_set': 150},
            # 부산 지역 부하
            {'name': 'BSN_Load', 'bus': 'BSN_EL', 'p_set': 120}
        ])
        
        # 선로 (지역간 연결)
        lines = pd.DataFrame([
            {'name': 'SEL_BSN_Line', 'bus0': 'SEL_EL', 'bus1': 'BSN_EL', 'x': 0.1, 'r': 0.01, 
             's_nom': 100, 's_nom_extendable': True, 's_nom_max': 500, 'carrier': 'AC'}
        ])
        
        # 수입/수출 링크
        links = []
        
        # 지역별 수입/수출 링크 추가
        for region in regions:
            # 수입 링크 (External_Grid -> 지역)
            links.append({
                'name': f"Import_{region}",
                'bus0': 'External_Grid',
                'bus1': f"{region}_EL",
                'p_nom': 200,
                'p_nom_extendable': True,
                'p_nom_min': 0,
                'p_nom_max': 1000,
                'efficiency': 1.0,
                'marginal_cost': 100,  # 높은 비용으로 내부 생산 우선
                'carrier': 'Link'
            })
            
            # 수출 링크 (지역 -> External_Grid)
            links.append({
                'name': f"Export_{region}",
                'bus0': f"{region}_EL",
                'bus1': 'External_Grid',
                'p_nom': 200,
                'p_nom_extendable': True,
                'p_nom_min': 0,
                'p_nom_max': 1000,
                'efficiency': 0.95,  # 송전 손실 반영
                'marginal_cost': -10,  # 수출은 음의 비용(=수익)으로 설정
                'carrier': 'Link'
            })
        
        links_df = pd.DataFrame(links)
        
        # 저장 장치
        stores = pd.DataFrame([
            {'name': 'SEL_Battery', 'bus': 'SEL_EL', 'e_nom': 50, 'e_nom_extendable': True, 
             'e_nom_max': 200, 'e_cyclic': True, 'standing_loss': 0.01, 
             'efficiency_store': 0.95, 'efficiency_dispatch': 0.95, 'carrier': 'battery'}
        ])
        
        # 시간 설정
        timeseries = pd.DataFrame([{
            'start_time': '2023-01-01 00:00:00',
            'end_time': '2023-01-02 00:00:00',  # 24시간
            'frequency': '1h'
        }])
        
        # 저장
        with pd.ExcelWriter(output_file) as writer:
            buses.to_excel(writer, sheet_name='buses', index=False)
            generators.to_excel(writer, sheet_name='generators', index=False)
            loads.to_excel(writer, sheet_name='loads', index=False)
            lines.to_excel(writer, sheet_name='lines', index=False)
            links_df.to_excel(writer, sheet_name='links', index=False)
            stores.to_excel(writer, sheet_name='stores', index=False)
            timeseries.to_excel(writer, sheet_name='timeseries', index=False)
        
        print(f"매우 단순한 모델이 '{output_file}'에 저장되었습니다.")
        print("이제 다음 명령어로 최적화를 실행해보세요:")
        print(f"python fix_optimize.py --input {output_file} --time day")
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    create_super_simple_model() 