import pandas as pd

df_gens = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

# carrier='coal'인 발전기 필터링
coal_gens = df_gens[df_gens['carrier'] == 'coal']

print("=" * 80)
print("carrier='coal'인 모든 발전기의 정확한 이름")
print("=" * 80)

# 발전기 이름에서 지역코드와 발전소명 추출
for idx, row in coal_gens.iterrows():
    name = row['name']
    name_lower = name.lower()
    
    # 지역코드 제거 (예: CND_, GND_ 등)
    if '_' in name:
        parts = name.split('_')
        plant_and_unit = '_'.join(parts[1:])  # 발전소명_호기
        plant_name = parts[1] if len(parts) > 1 else ''
    else:
        plant_name = name
    
    print(f"{name:45s} | 발전소명: {plant_name:20s} | lower: {name_lower}")

print("\n" + "=" * 80)
print("'기타'로 분류될 가능성이 있는 발전기 (발전소명 추출)")
print("=" * 80)

# 알려진 석탄발전소 목록
known_plants = ['당진', '영흥', '하동', '삼척', '태안', '보령', '삼천포', '호남', '동해', 
                '보스톤', '통영', '강릉안인', '북평', '신보령']

for idx, row in coal_gens.iterrows():
    name = row['name']
    name_lower = name.lower()
    
    # 발전소명 추출
    if '_' in name:
        parts = name.split('_')
        plant_name = parts[1] if len(parts) > 1 else ''
    else:
        plant_name = name
    
    # 알려진 발전소 목록에 있는지 확인
    found = False
    for known in known_plants:
        if known in name:
            found = True
            break
    
    if not found and 'coal' not in name_lower:
        print(f"[미매칭] {name:45s} | 발전소: {plant_name}")
        print(f"         원본 바이트: {plant_name.encode('utf-8')}")
        print(f"         16진수: {plant_name.encode('utf-8').hex()}")
