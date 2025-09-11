import pandas as pd

def check_excel_data():
    """Excel 파일의 실제 데이터 확인"""
    try:
        # Excel 파일 읽기
        xls = pd.ExcelFile("integrated_input_data.xlsx")
        
        print("=== Excel 파일 시트 목록 ===")
        print(xls.sheet_names)
        
        # buses 시트 확인
        print("\n=== buses 시트 데이터 ===")
        buses_df = pd.read_excel("integrated_input_data.xlsx", sheet_name="buses")
        print("버스 목록:")
        print(buses_df.to_string())
        
        # links 시트 확인
        print("\n=== links 시트 데이터 ===")
        links_df = pd.read_excel("integrated_input_data.xlsx", sheet_name="links")
        print("Links 컬럼:")
        print(list(links_df.columns))
        print("\nLinks 데이터:")
        print(links_df.to_string())
        
        # generators 시트 확인
        print("\n=== generators 시트 데이터 ===")
        generators_df = pd.read_excel("integrated_input_data.xlsx", sheet_name="generators")
        print("발전기 목록:")
        print(generators_df.to_string())
        
        # stores 시트 확인
        print("\n=== stores 시트 데이터 ===")
        stores_df = pd.read_excel("integrated_input_data.xlsx", sheet_name="stores")
        print("저장장치 목록:")
        print(stores_df.to_string())
        
        # loads 시트 확인
        print("\n=== loads 시트 데이터 ===")
        loads_df = pd.read_excel("integrated_input_data.xlsx", sheet_name="loads")
        print("부하 목록:")
        print(loads_df.to_string())
        
        return True
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
        return False

if __name__ == "__main__":
    check_excel_data() 