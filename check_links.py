import pandas as pd, warnings
warnings.filterwarnings('ignore')

print("=== integrated_input_data.xlsx 링크 확인 ===")
xl = pd.ExcelFile('integrated_input_data.xlsx')
print(f"시트: {xl.sheet_names}")

if 'links' in xl.sheet_names:
    df = xl.parse('links')
    print(f"\n링크 컬럼: {list(df.columns)}")
    print(f"\np_nom_extendable 현황:")
    if 'p_nom_extendable' in df.columns:
        cols = [c for c in ['name','p_nom_extendable','p_nom','p_nom_max','bus0','bus1'] if c in df.columns]
        print(df[cols].to_string())
    else:
        print("p_nom_extendable 컬럼 없음!")
        print(df[['name','p_nom']].head(20).to_string())

print("\n\n=== interface.xlsx 연도별_파라미터 시트 확인 ===")
xl2 = pd.ExcelFile('interface.xlsx')
if '연도별_파라미터' in xl2.sheet_names:
    df2 = xl2.parse('연도별_파라미터')
    print(df2.to_string())
else:
    print("연도별_파라미터 시트 없음!")

print("\n\n=== 2030 결과 링크 확인 ===")
import os
base = 'C:/Users/Hyodong.Moon/Desktop/HDMOON/python workplace/PyPSA_GESI_Test/results_multi/20260609_105051/2030'
files = os.listdir(base)
# nc 파일이 있으면 pypsa로 읽기
nc_files = [f for f in files if f.endswith('.nc')]
if nc_files:
    try:
        import pypsa
        n = pypsa.Network()
        n.import_from_netcdf(f'{base}/{nc_files[0]}')
        print("\n링크 p_nom_extendable:")
        print(n.links[['carrier','p_nom','p_nom_opt','p_nom_extendable']].to_string() if 'p_nom_opt' in n.links.columns else n.links[['carrier','p_nom','p_nom_extendable']].to_string())
    except Exception as e:
        print(f"오류: {e}")
