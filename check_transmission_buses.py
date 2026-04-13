import pandas as pd

print('='*70)
print('물리적 송전망 버스 확인')
print('='*70)

all_buses = pd.read_excel('interface.xlsx', sheet_name='all_buses')
virtual_buses = pd.read_excel('interface.xlsx', sheet_name='virtual_buses')

print(f'\n[current model]')
print(f'  all_buses (물리적 송전망): {len(all_buses)}개')
print(f'  virtual_buses (가상 버스): {len(virtual_buses)}개')
print(f'  total: {len(all_buses) + len(virtual_buses)}개')

print(f'\n[all_buses 샘플]')
print(all_buses.head(10).to_string())

print(f'\n[virtual_buses 샘플]')
print(virtual_buses.head(10).to_string())

print('\n' + '='*70)
