import sys
import os
sys.path.insert(0, os.path.dirname(__file__))

from PyPSA_GUI import read_input_data, create_network
import pandas as pd

print("=" * 80)
print("석탄발전기 세분화 테스트")
print("=" * 80)

# 데이터 로드
print("\n1. 데이터 로딩 중...")
input_data = read_input_data('integrated_input_data.xlsx')

# 발전기 확인
df_gens = input_data['generators']
coal_gens = df_gens[df_gens['name'].str.contains('당진|영흥|하동|삼척|태안|보령|삼천포|호남|동해', na=False)]

print(f"\n2. 세분화된 석탄발전기 확인:")
print(f"   개수: {len(coal_gens)}")
print(f"   총 용량: {coal_gens['p_nom'].sum():.1f} MW")
print(f"   Carrier: {coal_gens['carrier'].unique()}")
print(f"   Committable: {coal_gens['committable'].unique()}")
print(f"   한계비용 범위: {coal_gens['marginal_cost'].min():.1f} ~ {coal_gens['marginal_cost'].max():.1f}")

# LNG 발전기 확인
lng_gens = df_gens[df_gens['name'].str.contains('LNG', na=False)]
print(f"\n3. LNG 발전기 확인:")
print(f"   개수: {len(lng_gens)}")
print(f"   Committable: {lng_gens['committable'].unique()}")
print(f"   한계비용 범위: {lng_gens['marginal_cost'].min():.1f} ~ {lng_gens['marginal_cost'].max():.1f}")

print(f"\n4. 경제성 비교:")
print(f"   석탄 평균 한계비용: {coal_gens['marginal_cost'].mean():.2f} 원/MWh")
print(f"   LNG 평균 한계비용: {lng_gens['marginal_cost'].mean():.2f} 원/MWh")
print(f"   → 석탄이 {'저렴' if coal_gens['marginal_cost'].mean() < lng_gens['marginal_cost'].mean() else '비쌈'}")

print(f"\n5. 결론:")
if coal_gens['committable'].all() == False and coal_gens['carrier'].all() == 'electricity':
    print("   [OK] 석탄발전기 설정이 정상입니다!")
    print("   - Carrier: electricity (AC 버스에 연결 가능)")
    print("   - Committable: False (유연한 dispatch)")
    print("   - 한계비용이 LNG보다 저렴하여 우선 선택될 것으로 예상됩니다.")
else:
    print("   [WARNING] 일부 설정에 문제가 있을 수 있습니다.")
    
print("\n" + "=" * 80)
