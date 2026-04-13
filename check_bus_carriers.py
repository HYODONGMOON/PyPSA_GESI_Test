"""
버스 carrier 확인
"""
import pandas as pd
import sys
import io

# 인코딩 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

file_path = r"C:\Users\Hyodong.Moon\Desktop\HDMOON\python workplace\PyPSA_GESI_Test\integrated_input_data.xlsx"

buses = pd.read_excel(file_path, sheet_name='buses')

print("\n" + "="*70)
print("버스 Carrier 분석")
print("="*70 + "\n")

print(f"총 버스 수: {len(buses)}\n")

# Carrier별 분포
print("Carrier별 버스 분포:")
carrier_counts = buses['carrier'].value_counts()
print(carrier_counts)

print("\n각 carrier별 샘플 버스:")
for carrier in carrier_counts.index[:10]:
    sample_buses = buses[buses['carrier'] == carrier]['name'].head(5).tolist()
    print(f"\n[{carrier}]: {sample_buses}")

# 전력 버스로 추정되는 것들 확인
print("\n\n전력 버스로 추정되는 버스들:")
el_buses = buses[buses['name'].str.contains('_EL', na=False)]
print(f"총 {len(el_buses)}개")
if len(el_buses) > 0:
    print("\n샘플 (carrier 확인):")
    print(el_buses[['name', 'carrier', 'v_nom']].head(20))

print("\n" + "="*70 + "\n")
