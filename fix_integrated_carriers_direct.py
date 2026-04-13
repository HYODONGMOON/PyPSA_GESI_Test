import pandas as pd
import shutil
from datetime import datetime

# 백업 생성
backup_name = f'integrated_input_data_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
shutil.copy2('integrated_input_data.xlsx', backup_name)
print(f'백업 생성됨: {backup_name}')

# 파일 읽기
df = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

print('\n=== 수정 전 ===')
problem_gens = df[df['name'].str.contains('Coal|Hydro', na=False)]
print(problem_gens[['name', 'carrier', 'p_nom']].to_string(index=False))

# carrier 수정 로직
def fix_carrier(row):
    name = str(row['name'])
    current_carrier = str(row['carrier'])
    
    # 이름 기반으로 올바른 carrier 추론
    if 'Coal' in name and current_carrier == 'electricity':
        return 'coal'
    elif 'Hydro' in name and current_carrier == 'electricity':
        return 'hydro'
    else:
        return row['carrier']

# carrier 수정
df['carrier'] = df.apply(fix_carrier, axis=1)

print('\n=== 수정 후 ===')
problem_gens_after = df[df['name'].str.contains('Coal|Hydro', na=False)]
print(problem_gens_after[['name', 'carrier', 'p_nom']].to_string(index=False))

# 저장
with pd.ExcelWriter('integrated_input_data.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='generators', index=False)

print(f'\n수정 완료! 백업 파일: {backup_name}')
