import pandas as pd
import sys

def check_excel_structure(file_path):
    """엑셀 파일의 구조 확인"""
    try:
        print(f"\n파일 '{file_path}' 구조 분석 중...\n")
        
        # 엑셀 파일 읽기
        xls = pd.ExcelFile(file_path)
        
        # 모든 시트 목록 출력
        print("=== 시트 목록 ===")
        for idx, sheet_name in enumerate(xls.sheet_names):
            print(f"{idx+1}. {sheet_name}")
        
        # 각 시트별 상세 정보
        print("\n=== 각 시트별 상세 정보 ===")
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"\n## 시트: {sheet_name}")
            print(f"행 수: {len(df)}")
            print(f"열 수: {len(df.columns)}")
            print("컬럼 목록:")
            for col in df.columns:
                print(f"  - {col} (타입: {df[col].dtype})")
            
            # 데이터 샘플 (처음 5행)
            if not df.empty:
                print("\n데이터 샘플 (처음 5행):")
                print(df.head().to_string())
        
        return True
    
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        return False

if __name__ == "__main__":
    # 기본 파일 경로
    file_path = "input_data.xlsx"
    
    # 명령행 인수가 제공된 경우 해당 파일 사용
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    
    check_excel_structure(file_path) 