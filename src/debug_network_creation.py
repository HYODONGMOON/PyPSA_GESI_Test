import pandas as pd
import pypsa

def debug_network_creation():
    """네트워크 생성 과정 디버깅"""
    print("=== 네트워크 생성 디버깅 ===\n")
    
    try:
        # 데이터 로드
        input_data = {}
        xls = pd.ExcelFile('integrated_input_data.xlsx')
        
        for sheet_name in xls.sheet_names:
            input_data[sheet_name] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet_name)
        
        # 네트워크 생성
        network = pypsa.Network()
        
        # carriers 추가
        carriers = ['AC', 'DC', 'electricity', 'coal', 'gas', 'nuclear', 'solar', 'wind', 'hydrogen', 'heat']
        for carrier in carriers:
            network.add("Carrier", name=carrier, co2_emissions=0)
        
        # 시간 설정
        if 'timeseries' in input_data:
            ts = input_data['timeseries'].iloc[0]
            snapshots = pd.date_range(
                start=ts['start_time'],
                end=ts['end_time'],
                freq=ts['frequency'],
                inclusive='left'
            )
            network.set_snapshots(snapshots)
        
        # 버스 추가
        print("1. 버스 추가:")
        added_buses = []
        if 'buses' in input_data:
            for _, bus in input_data['buses'].iterrows():
                bus_name = str(bus['name'])
                network.add("Bus",
                          name=bus_name,
                          v_nom=float(bus['v_nom']),
                          carrier=str(bus['carrier']))
                added_buses.append(bus_name)
                print(f"   버스 추가: {bus_name}")
        
        print(f"\n총 {len(added_buses)}개 버스 추가됨")
        print(f"추가된 버스 목록: {sorted(added_buses)}")
        
        # 선로 연결 분석
        print("\n2. 선로 연결 분석:")
        if 'lines' in input_data:
            lines_df = input_data['lines']
            print(f"총 {len(lines_df)}개 선로 데이터")
            
            valid_lines = 0
            invalid_lines = 0
            
            for idx, line in lines_df.iterrows():
                line_name = str(line['name'])
                bus0_name = str(line['bus0'])
                bus1_name = str(line['bus1'])
                
                print(f"\n선로 {line_name}:")
                print(f"   bus0: {bus0_name}")
                print(f"   bus1: {bus1_name}")
                print(f"   bus0 존재: {bus0_name in added_buses}")
                print(f"   bus1 존재: {bus1_name in added_buses}")
                
                if bus0_name in added_buses and bus1_name in added_buses:
                    valid_lines += 1
                    print(f"   상태: 유효")
                    
                    # 실제로 선로 추가 시도
                    try:
                        network.add("Line",
                                  name=line_name,
                                  bus0=bus0_name,
                                  bus1=bus1_name,
                                  s_nom=float(line['s_nom']) if pd.notna(line['s_nom']) else 1000.0,
                                  x=float(line['x']) if pd.notna(line['x']) else 0.1,
                                  r=float(line['r']) if pd.notna(line['r']) else 0.01)
                        print(f"   선로 추가 성공!")
                    except Exception as e:
                        print(f"   선로 추가 실패: {str(e)}")
                        invalid_lines += 1
                        valid_lines -= 1
                else:
                    invalid_lines += 1
                    print(f"   상태: 유효하지 않음")
                    
                    # 누락된 버스 찾기
                    if bus0_name not in added_buses:
                        print(f"   누락된 버스: {bus0_name}")
                        # 비슷한 이름의 버스 찾기
                        similar_buses = [b for b in added_buses if bus0_name.split('_')[0] in b]
                        if similar_buses:
                            print(f"   유사한 버스들: {similar_buses}")
                    
                    if bus1_name not in added_buses:
                        print(f"   누락된 버스: {bus1_name}")
                        # 비슷한 이름의 버스 찾기
                        similar_buses = [b for b in added_buses if bus1_name.split('_')[0] in b]
                        if similar_buses:
                            print(f"   유사한 버스들: {similar_buses}")
            
            print(f"\n유효한 선로: {valid_lines}개")
            print(f"유효하지 않은 선로: {invalid_lines}개")
            
            # 네트워크에 추가된 선로 확인
            print(f"네트워크에 실제 추가된 선로: {len(network.lines)}개")
            if len(network.lines) > 0:
                print("추가된 선로 목록:")
                for line_name in network.lines.index:
                    print(f"   - {line_name}")
        
        return True
        
    except Exception as e:
        print(f"디버깅 중 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    debug_network_creation() 