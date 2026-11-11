import pandas as pd, warnings, os
warnings.filterwarnings('ignore')

base = 'C:/Users/Hyodong.Moon/Desktop/HDMOON/python workplace/PyPSA_GESI_Test/results_multi/20260609_105051'

for year, co2_limit in [('2030', 145.9e6), ('2035', 115.9e6)]:
    folder = f'{base}/{year}'
    files = os.listdir(folder)
    csv_fe = [f for f in files if 'final_energy_supply.csv' in f and 'region' not in f][0]
    df = pd.read_csv(f'{folder}/{csv_fe}')

    elec = df[df['final_energy'] == '전력']
    heat = df[df['final_energy'] == '열']

    lng_mwh    = elec[elec['technology'] == 'LNG']['supply_MWh'].sum()
    coal_mwh   = elec[elec['technology'] == '석탄']['supply_MWh'].sum()
    chp_mwh    = elec[elec['technology'] == 'CHP']['supply_MWh'].sum()
    solar_mwh  = elec[elec['technology'] == '태양광']['supply_MWh'].sum()
    wind_mwh   = elec[elec['technology'] == '풍력']['supply_MWh'].sum()
    nuke_mwh   = elec[elec['technology'] == '원자력']['supply_MWh'].sum()
    total_co2  = df['emissions_tCO2'].sum()
    elec_co2   = elec['emissions_tCO2'].sum()
    heat_co2   = heat['emissions_tCO2'].sum()

    lng_co2   = elec[elec['technology'] == 'LNG']['emissions_tCO2'].sum()
    coal_co2  = elec[elec['technology'] == '석탄']['emissions_tCO2'].sum()
    chp_elco2 = elec[elec['technology'] == 'CHP']['emissions_tCO2'].sum()
    chp_htco2 = heat[heat['technology'] == 'CHP']['emissions_tCO2'].sum()

    slack = co2_limit - total_co2

    print(f'\n{"="*60}')
    print(f'  {year}년  (CO2 제약: {co2_limit/1e6:.1f} MtCO2)')
    print(f'{"="*60}')
    print(f'  [발전량]')
    print(f'    LNG 발전:   {lng_mwh/1e9:.2f} TWh    CO2: {lng_co2/1e6:.2f} MtCO2')
    print(f'    석탄 발전:  {coal_mwh/1e9:.2f} TWh    CO2: {coal_co2/1e6:.2f} MtCO2')
    print(f'    CHP(전력):  {chp_mwh/1e9:.2f} TWh    CO2: {chp_elco2/1e6:.2f} MtCO2')
    print(f'    태양광:     {solar_mwh/1e9:.2f} TWh')
    print(f'    풍력:       {wind_mwh/1e9:.2f} TWh')
    print(f'    원자력:     {nuke_mwh/1e9:.2f} TWh')
    print(f'  [CO2 배출]')
    print(f'    전력분:     {elec_co2/1e6:.2f} MtCO2')
    print(f'    열(CHP):    {chp_htco2/1e6:.2f} MtCO2')
    print(f'    총계:       {total_co2/1e6:.3f} MtCO2')
    print(f'    CO2 여유:   {slack/1e6:.3f} MtCO2  --> {"[구속 binding]" if abs(slack) < 0.5e6 else "[비구속 non-binding]"}')
    print()
    print(f'  [석탄 배제 분석]')
    coal_co2_per_mwhe = 0.3431 / 0.44
    gas_co2_per_mwhe  = 0.1749 / 0.50
    if slack > 0:
        max_coal_twh = slack / ((coal_co2_per_mwhe - gas_co2_per_mwhe) * 1e6)
        print(f'    남은 CO2 여유분 {slack/1e6:.2f} MtCO2로 최대 가능 석탄: {max_coal_twh:.1f} TWh')
    else:
        print(f'    CO2 여유 없음: LNG만으로 이미 제약 초과 불가 상태 → 석탄 사용 불가')
