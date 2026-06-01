import re

content = open('PyPSA_GUI.py', encoding='utf-8').read()
lines = content.split('\n')

# 부하 관련 키워드 탐색
keywords = ['p_set', 'add_load', 'load_pattern', 'Demand_EL', 'n.loads', 
            'hourly_load', 'load_scale', 'annual_demand', 'loads_t']

found = {}
for i, line in enumerate(lines, 1):
    for kw in keywords:
        if kw.lower() in line.lower():
            if kw not in found:
                found[kw] = []
            found[kw].append((i, line.strip()[:120]))

for kw, hits in found.items():
    print(f'\n=== "{kw}" ({len(hits)}건) ===')
    for lineno, txt in hits[:8]:
        print(f'  L{lineno}: {txt}')

# 부하 추가 섹션 탐색
print('\n\n=== "=== Loads" 또는 "부하" 관련 섹션 ===')
for i, line in enumerate(lines, 1):
    if ('load' in line.lower() or '부하' in line) and ('====' in line or 'def ' in line or '시작' in line):
        print(f'  L{i}: {line.strip()[:120]}')
