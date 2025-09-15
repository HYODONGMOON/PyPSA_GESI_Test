import pandas as pd

def remove_backup_generators():
    """백업 LNG 발전기를 제거하는 함수"""
    
    print("=== 백업 LNG 발전기 제거 ===")
    
    try:
        # 발전기 데이터 로드
        generators_df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')
        print(f"현재 발전기 수: {len(generators_df)}")
        print("발전기 컬럼:", list(generators_df.columns))
        
        # 백업 발전기 식별
        backup_gens = generators_df[generators_df['name'].str.contains('Backup', na=False)]
        print(f"\n제거할 백업 발전기: {len(backup_gens)}개")
        
        if len(backup_gens) > 0:
            print("백업 발전기 목록:")
            total_backup_capacity = 0
            for _, gen in backup_gens.iterrows():
                name = gen['name']
                capacity = gen.get('p_nom', 0)
                total_backup_capacity += capacity
                print(f"- {name}: {capacity:.0f} MW")
            
            print(f"\n총 제거될 백업 용량: {total_backup_capacity:.0f} MW")
            
            # 백업 발전기 제거
            generators_cleaned = generators_df[~generators_df['name'].str.contains('Backup', na=False)]
            print(f"정리 후 발전기 수: {len(generators_cleaned)}개")
            
            # 다른 시트들 로드
            input_data = {}
            with pd.ExcelFile('integrated_input_data.xlsx') as xls:
                for sheet in xls.sheet_names:
                    if sheet != 'generators':
                        input_data[sheet] = pd.read_excel('integrated_input_data.xlsx', sheet_name=sheet)
            
            # 수정된 파일 저장
            with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl') as writer:
                generators_cleaned.to_excel(writer, sheet_name='generators', index=False)
                for sheet_name, df in input_data.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print("\n✅ 백업 LNG 발전기 제거 완료")
            print("✅ 이제 CHP가 우선적으로 사용될 것입니다")
            
            return True
            
        else:
            print("제거할 백업 발전기가 없습니다")
            return True
            
    except Exception as e:
        print(f"오류 발생: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    remove_backup_generators() 