import pandas as pd

# HP 링크들의 p_nom 값 확인
df = pd.read_excel('integrated_input_data.xlsx', sheet_name='links')
hp_links = df[df['name'].str.contains('HP', na=False)]

print('=== HP 링크들의 p_nom 값 ===')
for _, link in hp_links.iterrows():
    print(f'{link["name"]}: p_nom={link["p_nom"]}')

print(f'\n총 HP 링크 개수: {len(hp_links)}개')
print(f'p_nom이 0인 HP 링크: {len(hp_links[hp_links["p_nom"] == 0])}개')
print(f'p_nom이 100인 HP 링크: {len(hp_links[hp_links["p_nom"] == 100])}개') 