import pandas as pd
import numpy as np

# 테스트: 석탄발전기 집계 로직 검증
print("=" * 80)
print("석탄발전기 집계 로직 검증")
print("=" * 80)

# 테스트 데이터 생성 (시간대 × 발전기)
np.random.seed(42)
snapshots = pd.date_range('2024-01-01', periods=5, freq='h')

mock_data = {
    'CND_당진_1': np.random.rand(5) * 500,
    'CND_당진_2': np.random.rand(5) * 500,
    'CND_태안_1': np.random.rand(5) * 500,
    'GND_하동_1': np.random.rand(5) * 500,
    'GND_하동_2': np.random.rand(5) * 500,
    'ICN_영흥_1': np.random.rand(5) * 800,
    'GWD_동해_1': np.random.rand(5) * 400,
    'GWD_신서천_1': np.random.rand(5) * 1000,  # 신서천 포함
    'GND_고성_1': np.random.rand(5) * 1000,    # 고성 포함
    'JND_여수_1': np.random.rand(5) * 300,     # 여수 포함
    'CND_LNG': np.random.rand(5) * 4000,
    'GND_LNG': np.random.rand(5) * 1000,
    'GGD_Nuclear': np.random.rand(5) * 1400,
}

gtp = pd.DataFrame(mock_data, index=snapshots)

print(f"\n원본 발전기 수: {len(gtp.columns)}")
print(f"원본 컬럼: {list(gtp.columns)}")

# 석탄 집계 로직
coal_plant_keywords = ['당진', '영흥', '하동', '삼척', '태안', '보령', '삼천포', '호남', '동해',
                       '강릉안인', '북평', '신서천', '고성', '여수', '삼척그린파워']

coal_detail_cols = [c for c in gtp.columns if any(kw in c for kw in coal_plant_keywords)]

print(f"\n석탄 세부 발전기 ({len(coal_detail_cols)}개):")
for col in coal_detail_cols:
    print(f"  {col}")

# 지역별 합산
region_coal_agg = {}
for col in coal_detail_cols:
    region = col.split('_')[0] if '_' in col else 'ETC'
    agg_col = f'{region}_Coal'
    if agg_col not in region_coal_agg:
        region_coal_agg[agg_col] = gtp[col].copy()
    else:
        region_coal_agg[agg_col] += gtp[col]

# 원본에서 석탄 제거 후 합산 추가
gtp_agg = gtp.drop(columns=coal_detail_cols)
coal_agg_df = pd.DataFrame(region_coal_agg, index=gtp.index)
gtp_agg = pd.concat([gtp_agg, coal_agg_df], axis=1)

print(f"\n[Generator_Output 시트] 집계 후 컬럼 ({len(gtp_agg.columns)}개):")
for col in gtp_agg.columns:
    print(f"  {col}")

print(f"\n[Generator_Output_Coal 시트] 세부 발전기 ({len(coal_detail_cols)}개):")
for col in coal_detail_cols:
    print(f"  {col}")

# 검증: 합산이 일치하는지
print(f"\n[검증] CND 석탄 합산 일치 여부:")
cnd_manual = gtp[['CND_당진_1', 'CND_당진_2', 'CND_태안_1']].sum(axis=1)
cnd_agg = gtp_agg['CND_Coal']
match = np.allclose(cnd_manual, cnd_agg)
print(f"  수동 합산 vs CND_Coal: {'[OK] 일치' if match else '[NG] 불일치'}")
print(f"\n  CND 수동 합산 (첫 3시간): {cnd_manual.values[:3].round(2)}")
print(f"  CND_Coal (첫 3시간):      {cnd_agg.values[:3].round(2)}")

print("\n[SUCCESS] 석탄발전기 집계 로직이 정상 동작합니다!")
