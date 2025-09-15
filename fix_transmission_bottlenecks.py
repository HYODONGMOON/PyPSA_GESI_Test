import pandas as pd
from datetime import datetime
import shutil

def fix_transmission_bottlenecks():
    print("=== 송전 병목 지역 용량 증설 ===")
    
    try:
        # 백업 생성
        backup_file = f"integrated_input_data_before_bottleneck_fix_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
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
        
        # 병목 지역별 증설 계획
        bottleneck_fixes = [
            {"regions": ["DGU", "GND"], "additional_capacity": 2000, "reason": "205.8% 이용률 → 1,786MW 부족"},
            {"regions": ["DJN", "SJN"], "additional_capacity": 1200, "reason": "178.1% 이용률 → 991MW 부족"},
            {"regions": ["JJD", "JND"], "additional_capacity": 500, "reason": "116.1% 이용률 → 334MW 부족"},
            {"regions": ["GND", "JND"], "additional_capacity": 400, "reason": "112.9% 이용률 → 258MW 부족"},
            {"regions": ["CND", "GGD"], "additional_capacity": 800, "reason": "110.3% 이용률 → 562MW 부족"}
        ]
        
        print("\n=== 송전선로 증설 계획 ===")
        new_lines_added = 0
        
        for fix in bottleneck_fixes:
            region1, region2 = fix["regions"]
            additional_capacity = fix["additional_capacity"]
            reason = fix["reason"]
            
            # 기존 연결에서 참조할 선로 찾기
            reference_line = None
            for _, line in lines_df.iterrows():
                bus0_region = line['bus0'].split('_')[0] if '_' in line['bus0'] else line['bus0'][:3]
                bus1_region = line['bus1'].split('_')[0] if '_' in line['bus1'] else line['bus1'][:3]
                
                if (bus0_region == region1 and bus1_region == region2) or (bus0_region == region2 and bus1_region == region1):
                    reference_line = line
                    break
            
            if reference_line is not None:
                # 새로운 증설 선로 생성
                new_line = reference_line.copy()
                new_line['name'] = f"{region1}_{region2}_증설_{additional_capacity}MW"
                new_line['s_nom'] = additional_capacity
                new_line['s_nom_extendable'] = False
                
                # DataFrame에 추가
                lines_df = pd.concat([lines_df, pd.DataFrame([new_line])], ignore_index=True)
                new_lines_added += 1
                
                print(f"✅ {region1}-{region2}: +{additional_capacity}MW 증설")
                print(f"   이유: {reason}")
                print(f"   새 선로명: {new_line['name']}")
                print()
            else:
                print(f"⚠️ {region1}-{region2}: 참조 선로를 찾을 수 없음")
        
        print(f"총 {new_lines_added}개 선로 증설 완료")
        print(f"수정 후 송전선로: {len(lines_df)}개")
        
        # 수정된 lines 시트 업데이트
        input_data['lines'] = lines_df
        
        # Excel 파일로 저장
        with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
            for sheet_name, df in input_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"\n✅ 송전 병목 해결 완료!")
        print(f"주요 병목 지역에 총 {sum([fix['additional_capacity'] for fix in bottleneck_fixes])}MW 용량이 추가되었습니다.")
        
        # 증설 결과 요약
        print(f"\n=== 증설 결과 요약 ===")
        for fix in bottleneck_fixes:
            region1, region2 = fix["regions"]
            capacity = fix["additional_capacity"]
            print(f"- {region1}-{region2}: +{capacity}MW")
        
        print(f"\n이제 PyPSA 분석을 다시 실행하면 infeasible이 해결될 것입니다!")
        
        return True
        
    except Exception as e:
        print(f"증설 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    fix_transmission_bottlenecks() 