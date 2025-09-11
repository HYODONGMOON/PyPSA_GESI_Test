import pandas as pd

try:
    # 통합 데이터 파일 읽기 (가장 최근 백업 사용)
    input_file = 'integrated_input_data_backup_20250519_134702.xlsx'
    print(f"파일: {input_file}")
    
    # 버스 정보 확인
    df = pd.read_excel(input_file, sheet_name='buses')
    print(f"버스 수: {len(df)}")
    
    # BSN_EL 버스 확인
    bsn_exists = 'BSN_EL' in df['name'].values
    print(f"BSN_EL 버스 존재 확인: {bsn_exists}")
    
    # BSN 관련 버스 출력
    bsn_buses = df[df['name'].str.contains('BSN')]
    if len(bsn_buses) > 0:
        print("\nBSN 관련 버스 정보:")
        print(bsn_buses.to_string(index=False))
    else:
        print("\n버스에 BSN 관련 항목이 없습니다.")
    
    # 라인 정보도 확인
    print("\n" + "="*50)
    print("라인 정보 확인")
    print("="*50)
    lines_df = pd.read_excel(input_file, sheet_name='lines')
    print(f"라인 수: {len(lines_df)}")
    
    # 라인 출력시 필요한 칼럼만 선택
    line_columns = ['name', 'bus0', 'bus1', 'carrier', 's_nom', 'v_nom', 'length', 'x', 'r']
    
    # BSN_GND 라인 확인
    bsn_gnd_line = lines_df[lines_df['name'] == 'BSN_GND']
    if len(bsn_gnd_line) > 0:
        print("\nBSN_GND 라인 정보:")
        print(bsn_gnd_line[line_columns].to_string(index=False))
    else:
        print("\nBSN_GND 라인이 없습니다.")
        
    # BSN 관련 라인 모두 확인
    bsn_lines = lines_df[lines_df['name'].str.contains('BSN')]
    if len(bsn_lines) > 0:
        print("\n모든 BSN 관련 라인 정보:")
        for idx, line in bsn_lines.iterrows():
            print(f"\n라인 {idx+1}:")
            for col in line_columns:
                if col in line:
                    print(f"  {col}: {line[col]}")
    else:
        print("\n라인에 BSN 관련 항목이 없습니다.")
        
except Exception as e:
    print(f"오류 발생: {e}") 