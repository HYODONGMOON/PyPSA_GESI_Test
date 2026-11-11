import openpyxl, pandas as pd

wb = openpyxl.load_workbook('interface.xlsx', data_only=True)

print("=== constraints 시트 ===")
if 'constraints' in wb.sheetnames:
    ws = wb['constraints']
    for row in ws.iter_rows(min_row=1, max_row=20, values_only=True):
        if any(v is not None for v in row):
            print(row)

print("\n=== integrated_input_data constraints ===")
try:
    df = pd.read_excel('integrated_input_data.xlsx', sheet_name='constraints')
    print(df.to_string())
except Exception as e:
    print(f"오류: {e}")

print("\n=== CO2 emissions 계산 ===")
# coal carrier: co2_emissions = 0.73 * 0.47 = 0.3431
# gas carrier: co2_emissions = 0.33 * 0.53 = 0.1749
coal_eff = 0.43  # Coal_DB 평균 효율
gas_eff = 0.50

coal_co2_per_mwhe = 0.3431 / coal_eff
gas_co2_per_mwhe = 0.1749 / gas_eff

print(f"석탄 CO2/MWh_e = {0.3431:.4f}/{coal_eff} = {coal_co2_per_mwhe:.3f} tCO2/MWh_e")
print(f"가스 CO2/MWh_e = {0.1749:.4f}/{gas_eff} = {gas_co2_per_mwhe:.3f} tCO2/MWh_e")

# CO2 제약 145.9 MtCO2 = 145,900,000 tCO2
co2_limit = 145.9e6  # tCO2
fossil_demand_twh = 200  # 추정 화석연료 수요 (TWh)
fossil_demand_mwh = fossil_demand_twh * 1e6  # MWh

all_gas_co2 = fossil_demand_mwh * gas_co2_per_mwhe  # tCO2
print(f"\n화석연료 수요 {fossil_demand_twh} TWh 전부 LNG: CO2 = {all_gas_co2/1e6:.1f} MtCO2")
print(f"CO2 제약 145.9 MtCO2 대비 여유: {(co2_limit - all_gas_co2)/1e6:.1f} MtCO2")

slack_mwh = (co2_limit - all_gas_co2) / (coal_co2_per_mwhe - gas_co2_per_mwhe)
print(f"최대 석탄 사용량: {slack_mwh/1e6:.1f} TWh (이후 제약 초과)")
print("=> 석탄이 LNG보다 싸므로 이만큼 석탄이 먼저 사용되어야 함!")
