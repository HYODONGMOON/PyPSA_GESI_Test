import pandas as pd
from datetime import datetime
import shutil

def consolidate_transmission_lines():
    print("=== 지역간 송전선로 통합 ===")
    
    try:
        # 백업 생성
        backup_file = f"integrated_input_data_before_line_consolidation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        shutil.copy2('integrated_input_data.xlsx', backup_file)
        print(f"백업 파일 생성: {backup_file}")
        
        # 모든 시트 로드
        input_data = {}
        with pd.ExcelFile('integrated_input_data.xlsx') as xls:
            for sheet in xls.sheet_names:
                input_data[sheet] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet)
        
        # Lines 시트 수정
        lines_df = input_data['lines'].copy()
        print(f"수정 전 송전선로: {len(lines_df)}개")
        
        # 지역간 연결 분석
        regional_connections = {}
        
        for _, line in lines_df.iterrows():
            bus0_region = line['bus0'].split('_')[0] if '_' in line['bus0'] else line['bus0'][:3]
            bus1_region = line['bus1'].split('_')[0] if '_' in line['bus1'] else line['bus1'][:3]
            
            if bus0_region != bus1_region:
                # 정규화된 연결명 (알파벳 순)
                connection = f"{bus0_region}-{bus1_region}" if bus0_region < bus1_region else f"{bus1_region}-{bus0_region}"
                
                if connection not in regional_connections:
                    regional_connections[connection] = {
                        'lines': [],
                        'total_capacity': 0,
                        'bus0': line['bus0'],
                        'bus1': line['bus1']
                    }
                
                regional_connections[connection]['lines'].append(line)
                regional_connections[connection]['total_capacity'] += line['s_nom']
        
        # 새로운 통합 송전선로 생성
        new_lines = []
        region_lines_to_remove = []
        
        print("\n=== 지역간 연결 통합 ===")
        for connection, data in regional_connections.items():
            if len(data['lines']) > 1:  # 2개 이상의 선로가 있는 경우만 통합
                # 기존 선로들을 제거 대상에 추가
                for old_line in data['lines']:
                    region_lines_to_remove.append(old_line['name'])
                
                # 새로운 통합 선로 생성
                new_line = {
                    'name': f"{connection.replace('-', '_')}_통합",
                    'bus0': data['bus0'],
                    'bus1': data['bus1'],
                    's_nom': data['total_capacity'],
                    'length': data['lines'][0]['length'],  # 첫 번째 선로의 길이 사용
                    'r': data['lines'][0]['r'],  # 첫 번째 선로의 저항 사용
                    'x': data['lines'][0]['x'],  # 첫 번째 선로의 리액턴스 사용
                    'b': 0.0,  # 기본값
                    's_nom_extendable': False
                }
                
                new_lines.append(new_line)
                
                print(f"{connection}: {len(data['lines'])}개 선로 → 1개 통합 선로 ({data['total_capacity']:.1f} MW)")
                for old_line in data['lines']:
                    print(f"  - 제거: {old_line['name']} ({old_line['s_nom']:.1f} MW)")
                print(f"  + 추가: {new_line['name']} ({new_line['s_nom']:.1f} MW)")
                print()
        
        # 기존 lines_df에서 지역간 연결 선로 제거
        lines_df_filtered = lines_df[~lines_df['name'].isin(region_lines_to_remove)].copy()
        
        # 새로운 통합 선로 추가
        if new_lines:
            new_lines_df = pd.DataFrame(new_lines)
            lines_df_consolidated = pd.concat([lines_df_filtered, new_lines_df], ignore_index=True)
        else:
            lines_df_consolidated = lines_df_filtered
        
        print(f"수정 후 송전선로: {len(lines_df_consolidated)}개")
        print(f"제거된 선로: {len(region_lines_to_remove)}개")
        print(f"추가된 통합 선로: {len(new_lines)}개")
        
        # 통합된 lines 시트 업데이트
        input_data['lines'] = lines_df_consolidated
        
        # Excel 파일로 저장
        with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
            for sheet_name, df in input_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n✅ 송전선로 통합 완료!")
        print(f"주요 지역간 연결이 단일 고용량 선로로 통합되었습니다.")
        
        # 통합 결과 요약
        print(f"\n=== 통합 결과 요약 ===")
        major_consolidations = [
            ("ICN-GGD", 5899.9),
            ("GGD-CND", 5438.0), 
            ("GND-DGU", 1688.0),
            ("GGD-SEL", 9154.0)
        ]
        
        for conn_name, total_capacity in major_consolidations:
            consolidated_line = lines_df_consolidated[
                lines_df_consolidated['name'].str.contains(conn_name.replace('-', '_'), na=False)
            ]
            if not consolidated_line.empty:
                print(f"✓ {conn_name}: {total_capacity:.1f} MW로 통합됨")
            else:
                print(f"⚠ {conn_name}: 통합되지 않음 (확인 필요)")
                
        return True
        
    except Exception as e:
        print(f"통합 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    consolidate_transmission_lines() 