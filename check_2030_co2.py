"""
2030년 분석 결과에서 실제 LNG/Coal 발전량과 CO2를 역산하여
왜 석탄이 사용되지 않는지 진단합니다.
"""
import sys, os
sys.path.insert(0, os.path.abspath('.'))

import pandas as pd
import warnings
warnings.filterwarnings('ignore')

# ── 결과 파일 찾기 ────────────────────────────────────────────────────────
result_files = [f for f in os.listdir('.') if f.startswith('results') and f.endswith('.xlsx')
                and not f.startswith('~')]
print(f"결과 파일: {result_files}")

for rf in result_files[:2]:
    try:
        print(f"\n{'='*60}")
        print(f"파일: {rf}")
        xl = pd.ExcelFile(rf)
        print(f"시트 목록: {xl.sheet_names[:15]}")

        # Generator_Output 시트
        for sheet in ['Generator_Output', 'Generator_Output_Coal', 'FinalEnergy_Supply',
                      'CO2_Emissions', 'CO2', 'Summary']:
            if sheet in xl.sheet_names:
                df = xl.parse(sheet, index_col=0)
                print(f"\n  [{sheet}]  shape={df.shape}")
                print(df.head(10).to_string())
    except Exception as e:
        print(f"  오류: {e}")

# ── CO2 한계비용 역산 ─────────────────────────────────────────────────────
print("\n\n=== CO2 제약 임계값 계산 ===")
coal_co2_per_mwhe = 0.3431 / 0.44   # 평균 효율 0.44
gas_co2_per_mwhe  = 0.1749 / 0.50   # 평균 효율 0.50
coal_mc = 114.31
gas_mc  = 226.02

# LNG가 석탄보다 유리해지는 CO2 shadow price (lambda)
# coal_mc + lambda * coal_co2 > gas_mc + lambda * gas_co2
# lambda * (coal_co2 - gas_co2) > gas_mc - coal_mc
lambda_threshold = (gas_mc - coal_mc) / (coal_co2_per_mwhe - gas_co2_per_mwhe)
print(f"석탄 CO2: {coal_co2_per_mwhe:.3f} tCO2/MWh_e")
print(f"LNG  CO2: {gas_co2_per_mwhe:.3f} tCO2/MWh_e")
print(f"LNG가 석탄보다 유리해지는 CO2 shadow price: {lambda_threshold:.1f} KRW/tCO2")
print(f"  (= {lambda_threshold/1300:.1f} USD/tCO2 @ 1,300 KRW/USD)")
print()
print("CO2 제약이 구속적(binding)이면:")
print("  - shadow price 높을수록 LNG 선호 (CO2/MWh 낮으므로)")
print(f"  - shadow price > {lambda_threshold:.0f} KRW/tCO2 이면 LNG 우선, 석탄 배제")
print()
co2_limit_2030 = 145_900_000  # tCO2
all_gas_co2_per_twh = gas_co2_per_mwhe * 1e6  # tCO2/TWh
max_fossil_twh_all_gas = co2_limit_2030 / all_gas_co2_per_twh
print(f"2030 CO2 한도: {co2_limit_2030/1e6:.1f} MtCO2")
print(f"전부 LNG 사용 시 최대 발전량: {max_fossil_twh_all_gas:.1f} TWh")
print(f"  -> 재생에너지가 이보다 적으면 CO2 제약이 구속적 (석탄 배제)")
