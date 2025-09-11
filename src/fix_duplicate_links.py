import pandas as pd
import os
import shutil
from datetime import datetime

def fix_duplicate_links():
    """중복된 링크 문제를 해결하는 스크립트"""
    
    # 파일 로드
    input_file = 'integrated_input_data.xlsx'
    output_file = 'integrated_input_data_fixed_links.xlsx'
    
    # 원본 파일 백업
    backup_file = f"integrated_input_data_backup_links_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    print(f"원본 파일 백업: {backup_file}")
    shutil.copy2(input_file, backup_file)
    
    # 엑셀 파일 로드
    print(f"파일 분석: {input_file}")
    
    # 각 시트 로드
    with pd.ExcelFile(input_file) as xls:
        buses_df = pd.read_excel(xls, 'buses')
        links_df = pd.read_excel(xls, 'links')
        loads_df = pd.read_excel(xls, 'loads')
        generators_df = pd.read_excel(xls, 'generators')
        lines_df = pd.read_excel(xls, 'lines')
        stores_df = pd.read_excel(xls, 'stores')
    
    # 중복 링크 확인
    print("\n중복 링크 확인:")
    
    # 링크 이름 중복 검사
    link_name_counts = links_df['name'].value_counts()
    duplicate_link_names = link_name_counts[link_name_counts > 1].index.tolist()
    
    print(f"중복된 링크 이름: {duplicate_link_names}")
    
    if duplicate_link_names:
        for link_name in duplicate_link_names:
            dups = links_df[links_df['name'] == link_name]
            print(f"\n링크 '{link_name}'의 중복 ({len(dups)}개):")
            
            for idx, link in dups.iterrows():
                print(f"  #{idx}: bus0={link['bus0']}, bus1={link['bus1']}, carrier={link['carrier']}")
        
        # 중복 링크 제거 (가장 최근에 추가된 링크만 유지)
        print("\n중복 링크 해결 방법:")
        
        # 방법 1: 원본 중복 링크 제거 (이름을 변경하여 모두 유지)
        new_links_df = links_df.copy()
        for link_name in duplicate_link_names:
            dups = new_links_df[new_links_df['name'] == link_name]
            
            # 첫 번째 항목은 원래 이름 유지, 나머지는 이름 변경
            for i, (idx, link) in enumerate(dups.iloc[1:].iterrows()):
                # 유효한 링크만 이름 변경 (실제 Hydrogen과 연결하는 링크)
                if link['bus1'] in ['SEL_Hydrogen', 'JBD_Hydrogen']:
                    new_name = f"{link_name}_Hydrogen"
                    print(f"  링크 #{idx}: '{link_name}' → '{new_name}'")
                    new_links_df.at[idx, 'name'] = new_name
        
        # 구성 요소 수 확인
        print(f"\n링크 수 (변경 전): {len(links_df)}")
        print(f"링크 수 (변경 후): {len(new_links_df)}")
        
        # 모든 데이터를 새 파일에 저장
        with pd.ExcelWriter(output_file) as writer:
            buses_df.to_excel(writer, sheet_name='buses', index=False)
            new_links_df.to_excel(writer, sheet_name='links', index=False)
            loads_df.to_excel(writer, sheet_name='loads', index=False)
            generators_df.to_excel(writer, sheet_name='generators', index=False)
            lines_df.to_excel(writer, sheet_name='lines', index=False)
            stores_df.to_excel(writer, sheet_name='stores', index=False)
        
        print(f"\n수정된 데이터가 {output_file}에 저장되었습니다.")
        print("이제 이 파일을 사용하여 네트워크 최적화를 시도해 보세요.")
    else:
        print("중복된 링크가 없습니다. 수정이 필요하지 않습니다.")

if __name__ == "__main__":
    fix_duplicate_links() 