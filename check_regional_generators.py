"""지역별 발전기 현황 확인"""
import pandas as pd

gen = pd.read_excel('integrated_input_data.xlsx', sheet_name='generators')

print("=" * 80)
print("지역별 발전기 현황")
print("=" * 80)

for region in ['SEL', 'JBD', 'DGU', 'SJN', 'JND', 'GBD']:
    print(f"\n[{region}]")
    region_gen = gen[gen['name'].str.startswith(region)]
    
    if len(region_gen) > 0:
        print(f"발전기 수: {len(region_gen)}개")
        print(f"총 용량: {region_gen['p_nom'].sum():,.0f} MW")
        print("\n상세:")
        for _, row in region_gen.iterrows():
            mcost = row.get('marginal_cost', 0)
            print(f"  - {row['name']:30s} | {row['carrier']:12s} | {row['p_nom']:>8,.0f} MW | {mcost:>8,.1f} 원/MWh")
    else:
        print("  발전기 없음")

print("\n" + "=" * 80)
print("슬랙 발전기 (slack, Fallback)")
print("=" * 80)

slack_gen = gen[gen['name'].str.contains('slack|Slack|Fallback', case=False, na=False)]
print(f"\n총 슬랙 발전기: {len(slack_gen)}개")
print(f"총 용량: {slack_gen['p_nom'].sum():,.0f} MW")

for region in ['SEL', 'JBD', 'DGU', 'SJN']:
    region_slack = slack_gen[slack_gen['name'].str.startswith(region)]
    if len(region_slack) > 0:
        print(f"\n[{region}] 슬랙 발전기: {len(region_slack)}개")
        for _, row in region_slack.iterrows():
            mcost = row.get('marginal_cost', 999)
            print(f"  - {row['name']:30s} | {row['p_nom']:>8,.0f} MW | {mcost:>8,.1f} 원/MWh")
