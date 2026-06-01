# PyPSA-GESI: 한국 에너지 시스템 최적화 모델

**녹색에너지전략연구소(GESI)** 에서 개발한 한국 전력·에너지 시스템 최적화 도구입니다.  
[PyPSA(Python for Power System Analysis)](https://pypsa.org/) 프레임워크 기반으로, 17개 광역 지자체를 노드로 설정한 **전국 단위 멀티-섹터 에너지 시스템 모델**입니다.

---

## 주요 특징

- **멀티-캐리어 에너지 모델링**: 전력(AC/DC), 열에너지, 수소, 전기차(EV) 부하를 동일 네트워크에서 통합 분석
- **17개 광역 지자체 노드**: 서울(SEL), 부산(BSN), 대구(DGU), 인천(ICN), 광주(GWJ), 대전(DJN), 울산(USN), 세종(SJG), 경기(GGD), 강원(GWD), 충북(CBD), 충남(CND), 전북(JBD), 전남(JND), 경북(GBD), 경남(GND), 제주(JJD)
- **송전망 모델링**: 지역 간 AC/DC 선로(HVDC 포함), 유연 운영(Flexible Line Operation) 기능
- **시나리오 기반 장기 분석**: `interface.xlsx`에서 연도별 수요·발전설비 시나리오를 정의하여 다년도 순차 최적화 실행
- **Unit Commitment(UC) 분석**: 석탄·가스 발전기의 주 단위 UC 분석 지원
- **풍부한 결과 시각화**: 지역별·계절별 발전량 스택, 송전 포화도 히트맵, 재생에너지 지도, 한국 지도 기반 시각화

---

## 모델 구조

```
PyPSA-GESI 모델
├── 입력 데이터
│   ├── integrated_input_data.xlsx   # 통합 입력 파일 (버스, 발전기, 선로, 수요 등)
│   └── interface.xlsx               # 시나리오 설정 (연도별 수요·용량·가격 등)
│
├── 핵심 모듈 (PyPSA_GUI.py)
│   ├── read_input_data()            # 엑셀 기반 입력 데이터 로드
│   ├── create_network()             # PyPSA 네트워크 객체 생성
│   ├── optimize_network()           # 선형/혼합정수 최적화 실행
│   ├── save_results()               # 결과 저장 (Excel, CSV, NetCDF)
│   ├── run_multi_year_sequence()    # 다년도 순차 시나리오 분석
│   └── create_visualizations()     # 시각화 차트 생성
│
├── 보조 모듈 (modules/)
│   ├── data_loader.py
│   ├── network_builder.py
│   ├── optimizer.py
│   ├── result_processor.py
│   └── visualization.py
│
└── 결과 (results/)
    ├── *.xlsx                       # 지역별·발전원별 발전량, 송전 현황 등
    ├── *.csv                        # 상세 시계열 데이터
    ├── *.nc                         # PyPSA 네트워크 전체 결과 (NetCDF)
    └── images/                      # 시각화 결과 이미지
```

---

## 에너지 캐리어 구성

| 캐리어 | 설명 | 포함 기술 |
|--------|------|-----------|
| `electricity` | 전력 (AC) | 석탄, LNG, 원자력, 태양광, 풍력, 수력, 양수 |
| `DC` | 직류 송전 | HVDC 연계선 |
| `heat` | 열에너지 | 열펌프(HP), CHP, 지역난방 |
| `hydrogen` | 수소 | 전해조(Electrolyzer), 수소 연료전지 |
| `EV` | 전기차 | V2G, 스마트 충전 부하 |

---

## 빠른 시작

### 1. 환경 설정

```bash
conda create -n pypsa_env python=3.10
conda activate pypsa_env
pip install -r requirements.txt
```

> CPLEX 또는 Gurobi 솔버 사용 시 별도 라이선스 설치 필요 (기본: GLPK/HiGHS 무료 솔버 지원)

### 2. 입력 데이터 준비

- `integrated_input_data.xlsx`: 버스, 발전기, 선로, 저장장치, 부하, 시계열 패턴을 시트별로 정의
- `interface.xlsx`: 연도별 에너지 수요 시나리오, 발전설비 용량 계획, 유연 선로 운영 설정

### 3. 단일 연도 분석 실행

```bash
python PyPSA_GUI.py
```

### 4. 다년도 시나리오 분석

```python
from PyPSA_GUI import run_multi_year_sequence, build_overrides_for_years

years = [2030, 2035, 2040, 2045, 2050]
overrides = build_overrides_for_years(years, base_input_file='integrated_input_data.xlsx')

results = run_multi_year_sequence(
    years=years,
    overrides_by_year=overrides,
    carryover=True,              # 이전 연도 설비 용량 인계
    results_root='results_multi'
)
```

---

## 주요 출력 결과

### Excel / CSV
| 파일명 패턴 | 내용 |
|-------------|------|
| `*_generator_output.csv` | 발전기별 시간별 발전량 |
| `*_regional_power_balance.csv` | 지역별 수급 균형 |
| `*_line_usage.csv` | 송전선로 흐름 및 포화도 |
| `*_final_energy_supply.csv` | 최종에너지 공급량 (섹터별) |
| `*_발전원별_발전량.csv` | 발전원별 연간 발전량 |
| `*_지역별_발전원별_발전량.csv` | 지역×발전원 교차 분석 |
| `*.xlsx` | 종합 분석 보고서 (다중 시트) |

### 시각화
| 차트 | 설명 |
|------|------|
| `viz_01_line_utilization` | 송전선로 활용률 분포 |
| `viz_02_re_vs_tx` | 재생에너지 비중 vs 송전 혼잡도 |
| `viz_03_seasonal_heatmap` | 계절별 시간대 발전량 히트맵 |
| `viz_04_regional_balance` | 지역별 에너지 수급 균형 |
| `viz_05_sector_coupling` | 섹터 커플링 현황 |
| `regional_analysis/` | 지역×계절별 상세 분석 (5종 차트) |
| `*_korea_transmission_map` | 한국 지도 기반 송전망 현황 |
| `*_re_korea_map` | 한국 지도 기반 재생에너지 분포 |

---

## 시나리오 분석 구조

`interface.xlsx`의 시트 구성:

| 시트명 | 내용 |
|--------|------|
| `인터페이스_1` | 선로별 유연 운영 설정 (s_nom_flex, 연간 허용 시간 등) |
| `시나리오_에너지수요` | 연도별·지역별·섹터별 에너지 수요 시나리오 |
| 지역 코드 시트 (예: `GGD`, `CND` ...) | 지역별 발전기·링크·부하 설정 |
| `load_patterns` | 지역별·섹터별 8760시간 부하 패턴 |
| `renewable_patterns` | 지역별 태양광·풍력 발전량 패턴 |

---

## 기술 스택

| 구분 | 사용 기술 |
|------|-----------|
| 최적화 프레임워크 | [PyPSA](https://pypsa.org/) >= 0.21 |
| 솔버 | CPLEX / Gurobi / HiGHS (GLPK) |
| 언어 | Python 3.10+ |
| 데이터 처리 | pandas, numpy, openpyxl |
| 시각화 | matplotlib, seaborn, plotly, folium |
| 지리 데이터 | geopandas, shapely, pyproj |
| 결과 저장 | NetCDF4 (`.nc`), Excel (`.xlsx`), CSV |

---

## 디렉터리 구조

```
PyPSA_GESI_Test/
├── PyPSA_GUI.py                     # 메인 실행 파일 (네트워크 생성·최적화·결과 저장)
├── integrated_input_data.xlsx       # 통합 입력 데이터
├── interface.xlsx                   # 시나리오 인터페이스 설정
├── requirements.txt                 # 패키지 의존성
├── README_CPLEX.md                  # CPLEX 솔버 설정 가이드
├── modules/                         # 보조 모듈
│   ├── data_loader.py
│   ├── network_builder.py
│   ├── optimizer.py
│   ├── result_processor.py
│   └── visualization.py
├── src/                             # 지도 시각화 등 추가 소스
│   └── korea_map.py
├── data/                            # 부하·재생에너지 시계열 원시 데이터
└── results/                         # 최적화 결과 (gitignore 처리)
```

---

## 라이선스 및 문의

- 이 모델은 **녹색에너지전략연구소(GESI)** 내부 연구 목적으로 개발되었습니다.
- 문의: [GESI 홈페이지](https://gesi.kr)

---

**버전**: 2.0.0  
**최종 업데이트**: 2026-06  
**개발 환경**: Python 3.10, PyPSA 0.21+, Windows 10/11
