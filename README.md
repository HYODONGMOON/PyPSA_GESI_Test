# PyPSA_KOREA_GESI 🇰🇷

한국 전력 시스템 최적화를 위한 PyPSA 기반 에너지 시스템 통합 분석 도구

## 📋 프로젝트 개요

PyPSA_KOREA_GESI는 한국의 전력 시스템을 모델링하고 최적화하기 위한 종합적인 도구입니다. PyPSA(Python for Power System Analysis) 프레임워크를 기반으로 하여 한국의 17개 광역시도별 에너지 시스템을 분석할 수 있습니다.

### 🎯 주요 기능

- **지역별 전력 시스템 모델링**: 17개 광역시도별 발전, 송전, 부하 시스템
- **재생에너지 통합 분석**: 태양광(PV), 풍력(WT) 발전 패턴 적용
- **송전선로 최적화**: 지역간 전력 조류 및 송전 용량 분석
- **에너지 저장 시스템**: ESS, 수소 저장 등 다양한 저장 기술 모델링
- **시각화**: 한국 지도 기반 결과 시각화 및 대시보드

## 🚀 빠른 시작

### 1. 환경 설정

```bash
# 저장소 클론
git clone https://github.com/your-username/PyPSA_KOREA_GESI.git
cd PyPSA_KOREA_GESI

# 가상환경 생성 (권장)
python -m venv pypsa_env
source pypsa_env/bin/activate  # Linux/Mac
# 또는
pypsa_env\Scripts\activate     # Windows

# 패키지 설치
pip install -r requirements.txt
```

### 2. 실행

```bash
# 메인 분석 실행
python PyPSA_GUI.py
```

## 📁 프로젝트 구조

```
PyPSA_KOREA_GESI/
├── PyPSA_GUI.py                    # 메인 실행 파일
├── integrated_input_data.xlsx      # 통합 입력 데이터
├── README.md
├── requirements.txt
├── .gitignore
├── src/                            # 소스 코드
│   ├── korea_map.py               # 한국 지도 시각화
│   ├── analyze_regional_results.py # 지역별 결과 분석
│   ├── PyPSA_HD_Regional.py       # 지역별 모델링
│   └── ...
├── data/                          # 입력 데이터
│   ├── integrated_input_data.xlsx
│   ├── regional_input_template.xlsx
│   └── map_data/                  # 한국 지도 데이터
└── results/                       # 결과 파일 (자동 생성)
```

## 🔧 주요 구성 요소

### 1. 네트워크 모델링
- **버스(Bus)**: 17개 광역시도별 AC, 수소, 열 버스
- **발전기(Generator)**: 원자력, 석탄, LNG, 태양광, 풍력, 수소 발전
- **부하(Load)**: 전력, 수소, 열 부하
- **저장장치(Store)**: 배터리, 수소 저장
- **송전선로(Line)**: 지역간 송전 연결

### 2. 데이터 구조
- **buses**: 버스 정보 (이름, 전압, 캐리어, 좌표)
- **generators**: 발전기 정보 (용량, 비용, 효율)
- **loads**: 부하 정보 (시간별 부하 패턴)
- **stores**: 저장장치 정보 (용량, 효율)
- **lines**: 송전선로 정보 (용량, 임피던스)
- **renewable_patterns**: 재생에너지 발전 패턴
- **load_patterns**: 지역별 부하 패턴

### 3. 최적화 엔진
- **CPLEX Solver**: 고성능 선형 최적화
- **병렬 처리**: 멀티코어 CPU 활용
- **제약 조건**: CO2 배출 제한, 용량 제약

## 📊 결과 분석

실행 후 `results/` 폴더에 다음 결과들이 생성됩니다:

### 📈 Excel 결과 파일
- `optimization_result_YYYYMMDD_HHMMSS.xlsx`: 종합 최적화 결과
- 시트별 상세 결과:
  - Generator_Output: 발전기별 시간별 출력
  - Line_Flow: 송전선로별 조류
  - Storage_Power: 저장장치 충방전
  - Bus_Info: 버스 정보
  - Summary: 최적화 요약

### 📊 시각화 결과
- `regional_energy_balance.png`: 지역별 에너지 밸런스
- `regional_renewable_ratio.png`: 지역별 재생에너지 비율
- `transmission_network_graph.png`: 송전망 네트워크
- `korea_map.html`: 인터랙티브 한국 지도
- `transmission_flow_map.html`: 송전 조류 지도

### 📋 CSV 데이터
- `generator_output.csv`: 발전기 출력 데이터
- `load.csv`: 부하 데이터
- `storage.csv`: 저장장치 데이터
- `line_usage.csv`: 송전선로 이용률

## ⚙️ 설정 및 커스터마이징

### 1. 시간 설정
`data/regional_time_settings.json`에서 분석 기간 설정:
```json
{
  "start_time": "2023-01-01 00:00:00",
  "end_time": "2023-12-31 23:00:00",
  "frequency": "1h"
}
```

### 2. 지역별 데이터 수정
`data/regional_input_template.xlsx`에서 지역별 파라미터 조정 가능

### 3. 재생에너지 패턴
`integrated_input_data.xlsx`의 `renewable_patterns` 시트에서 시간별 발전 패턴 설정

## 🛠️ 기술 스택

- **Python 3.8+**
- **PyPSA**: 전력 시스템 분석 프레임워크
- **CPLEX**: 최적화 솔버
- **Pandas**: 데이터 처리
- **Matplotlib/Plotly**: 시각화
- **Geopandas**: 지리 데이터 처리
- **NetworkX**: 네트워크 분석

## 📋 요구사항

### 필수 패키지
```
pypsa>=0.21.0
pandas>=1.3.0
numpy>=1.21.0
matplotlib>=3.5.0
plotly>=5.0.0
geopandas>=0.10.0
networkx>=2.6.0
openpyxl>=3.0.0
```

### 솔버
- **CPLEX**: 상용 최적화 솔버 (권장)
- **Gurobi**: 대안 상용 솔버
- **HiGHS**: 오픈소스 솔버 (기본)

## 🔍 사용 예시

### 기본 실행
```python
from PyPSA_GUI import *

# 데이터 로드
input_data = read_input_data('integrated_input_data.xlsx')

# 네트워크 생성
network = create_network(input_data)

# 최적화 실행
optimize_network(network)

# 결과 저장
save_results(network)
```

### 지역별 분석
```python
from src.analyze_regional_results import analyze_regional_results

# 지역별 상세 분석
analyze_regional_results(network, 'results/', '20231201_120000')
```

## 🤝 기여하기

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📄 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 자세한 내용은 `LICENSE` 파일을 참조하세요.

## 📞 연락처

- **프로젝트 링크**: [https://github.com/your-username/PyPSA_KOREA_GESI](https://github.com/your-username/PyPSA_KOREA_GESI)
- **이슈 리포트**: [Issues](https://github.com/your-username/PyPSA_KOREA_GESI/issues)

## 🙏 감사의 말

- [PyPSA](https://pypsa.org/) 개발팀
- 한국 전력 시스템 데이터 제공 기관들
- 오픈소스 커뮤니티

---

**PyPSA_KOREA_GESI**로 한국의 지속가능한 에너지 미래를 설계해보세요! 🌱⚡ 