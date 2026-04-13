# 단계 1: 코드 분석 결과

## ✅ 완료 시각: 2026-01-29

---

## 📊 PyPSA_GUI.py 구조 분석

### 🔍 **주요 함수 목록 (총 53개)**

#### 1. 데이터 관련
- `read_input_data(input_file)` - Excel에서 데이터 로딩
- `validate_input_data(input_data)` - 입력 데이터 검증
- `standardize_bus_names_in_input(input_data)` - 버스 이름 표준화

#### 2. 네트워크 생성
- `create_network(input_data)` - PyPSA 네트워크 생성
  - carriers 정의
  - snapshots 설정 (8760시간)
  - buses, generators, loads, stores, links, lines 추가

#### 3. 최적화
- `optimize_network(network)` - 네트워크 최적화 실행
  - CPLEX 솔버 설정
  - network.optimize() 호출

#### 4. 결과 처리
- `extract_results(network)` - 주요 결과 추출
- `save_results(network)` - Excel 파일로 저장
- `create_visualizations(network)` - 시각화 생성

#### 5. 메인 실행
- `main()` - 전체 프로세스 오케스트레이션

---

## 🔄 **현재 실행 흐름**

```python
main()
├── ensure_integrated_input()                # interface.xlsx → integrated_input_data.xlsx
├── read_input_data(INPUT_FILE)              # 데이터 로딩
│   ├── buses, generators, loads, stores
│   ├── links, lines
│   ├── timeseries (8760시간)
│   ├── renewable_patterns
│   └── load_patterns
├── standardize_bus_names_in_input()         # 버스 이름 정리
├── create_network(input_data)               # 네트워크 생성
│   ├── carriers 추가
│   ├── snapshots 설정 (8760시간)
│   ├── buses 추가 (503개)
│   ├── generators 추가 (62개)
│   ├── loads 추가 (시계열 데이터 연결)
│   ├── stores 추가 (76개)
│   ├── links 추가 (68개)
│   └── lines 추가 (1473개)
├── optimize_network(network)                # ⬅️ 여기를 Rolling Horizon으로 대체!
│   ├── CPLEX 설정 (40GB 메모리)
│   ├── network.optimize()
│   └── 결과 반환
└── save_results(network)                    # 결과 저장
    ├── Excel 파일
    ├── 시각화
    └── 지도
```

---

## 🎯 **Rolling Horizon 삽입 위치**

### **옵션 1: 별도 함수 생성 (권장)** ✅

```python
def rolling_horizon_optimize(network, input_data, num_segments=12):
    """
    Rolling Horizon 최적화
    
    Args:
        network: 기본 네트워크 (템플릿)
        input_data: 전체 입력 데이터
        num_segments: 구간 수 (기본 12개월)
    
    Returns:
        통합된 네트워크 결과
    """
    # 구간별 최적화 로직
    pass

# main() 수정
def main():
    # ... 기존 코드 ...
    network = create_network(input_data)
    
    # Rolling Horizon 사용 여부 선택
    use_rolling_horizon = True  # GUI에서 설정 가능
    
    if use_rolling_horizon:
        network = rolling_horizon_optimize(network, input_data, num_segments=12)
    else:
        optimize_network(network)
    
    save_results(network)
```

**장점:**
- 기존 코드 영향 최소화
- 선택적 사용 가능
- 테스트 용이

---

## 📐 **데이터 구조 분석**

### 1. input_data 딕셔너리
```python
input_data = {
    'buses': DataFrame,           # 503개
    'generators': DataFrame,      # 62개
    'loads': DataFrame,           # 68개
    'stores': DataFrame,          # 76개
    'links': DataFrame,           # 68개
    'lines': DataFrame,           # 3524개 (1473개 사용)
    'timeseries': DataFrame,      # 시작/종료 시간
    'renewable_patterns': DataFrame,  # 8760 × N
    'load_patterns': DataFrame    # 8760 × M
}
```

### 2. 시계열 데이터 구조
```python
# renewable_patterns
columns: ['solar_pattern', 'wind_pattern', ...]
index: 0-8759 (8760 rows)

# load_patterns  
columns: ['BSN_EL', 'CBD_EL', 'CND_EL', ...]
index: 0-8759 (8760 rows)
```

### 3. expandable 컴포넌트
```python
# generators
columns: [..., 'p_nom', 'p_nom_extendable', 'p_nom_min', 'p_nom_max', ...]

# stores
columns: [..., 'e_nom', 'e_nom_extendable', 'e_nom_min', 'e_nom_max', ...]

# lines
columns: [..., 's_nom', 's_nom_extendable', 's_nom_min', 's_nom_max', ...]

# links
columns: [..., 'p_nom', 'p_nom_extendable', 'p_nom_min', 'p_nom_max', ...]
```

### 4. 저장소 초기 상태
```python
# stores
columns: [..., 'e_initial', 'e_cyclic', ...]

# interface.xlsx에서 설정 (예: 0.5 = 50%)
```

### 5. CO2 제약
```python
# 전역 제약으로 설정됨
# carriers에 co2_emissions 정의
# GlobalConstraint으로 총량 제한
```

---

## 🔧 **Rolling Horizon 구현을 위한 필요 작업**

### 1. 시계열 데이터 슬라이싱
```python
# 구간별 시간 인덱스 계산
segment_hours = 730  # 약 1개월
start_idx = segment_id * segment_hours
end_idx = start_idx + segment_hours

# renewable_patterns 슬라이싱
segment_re_patterns = renewable_patterns.iloc[start_idx:end_idx]

# load_patterns 슬라이싱
segment_load_patterns = load_patterns.iloc[start_idx:end_idx]
```

### 2. 구간별 네트워크 생성
```python
# 기본 네트워크 복사
network_segment = network.copy()

# snapshots 재설정
segment_snapshots = network.snapshots[start_idx:end_idx]
network_segment.set_snapshots(segment_snapshots)

# 시계열 데이터 업데이트
for gen in network_segment.generators.index:
    if gen has renewable pattern:
        network_segment.generators_t.p_max_pu[gen] = segment_re_patterns[pattern_col]
```

### 3. 설비 용량 전달
```python
# 이전 구간 결과
prev_capacities = {
    'generators': {gen: network_prev.generators.loc[gen, 'p_nom_opt']},
    'stores': {store: network_prev.stores.loc[store, 'e_nom_opt']},
    'lines': {line: network_prev.lines.loc[line, 's_nom_opt']},
    'links': {link: network_prev.links.loc[link, 'p_nom_opt']}
}

# 다음 구간에 적용
network_next.generators.loc[gen, 'p_nom'] = prev_capacities['generators'][gen]
```

### 4. 저장소 상태 전달
```python
# 이전 구간 마지막 SOC
prev_soc = network_prev.stores_t.e.iloc[-1]

# 다음 구간 초기 SOC
for store in network_next.stores.index:
    network_next.stores.loc[store, 'e_initial'] = prev_soc[store]
```

### 5. CO2 제약 할당
```python
# 연간 총 제약
annual_co2_limit = get_annual_co2_limit(input_data)

# 구간별 수요
segment_demand = network_segment.loads_t.p_set.sum().sum()
total_demand = sum(all_segment_demands)

# 비례 할당
segment_co2_limit = annual_co2_limit * (segment_demand / total_demand)

# 제약 설정
network_segment.add("GlobalConstraint",
                   "co2_limit",
                   type="primary_energy",
                   carrier_attribute="co2_emissions",
                   sense="<=",
                   constant=segment_co2_limit)
```

---

## 📍 **주요 발견사항**

### ✅ 좋은 점
1. **모듈화**: 함수들이 명확히 분리되어 있음
2. **데이터 구조**: input_data 딕셔너리가 잘 정의됨
3. **시계열 데이터**: renewable_patterns, load_patterns가 분리되어 있어 슬라이싱 용이
4. **확장성**: expandable 옵션이 이미 구현되어 있음

### ⚠️ 주의사항
1. **network.copy()**: 네트워크 복사 시 깊은 복사 필요
2. **snapshots 재설정**: 구간별로 시간 인덱스 정확히 맞춰야 함
3. **시계열 데이터 인덱스**: pandas iloc vs loc 주의
4. **메모리 관리**: 12개 구간 결과를 모두 메모리에 보관하지 말고 순차 저장

---

## 🎯 **다음 단계 (단계 2) 준비사항**

### 단계 2에서 할 일:
1. `rolling_horizon_optimize()` 함수 뼈대 작성
2. 12개월 구간 분할 로직
3. 구간별 시간 인덱스 계산
4. 기본 구조 테스트 (실제 최적화 없이)

### 필요한 정보:
- [x] main() 함수 구조 파악
- [x] create_network() 구조 파악
- [x] optimize_network() 구조 파악
- [x] 데이터 구조 파악
- [x] 시계열 데이터 처리 방법 파악

---

## ✅ **단계 1 완료!**

**다음 단계로 진행 준비 완료**

**예상 작업 시간:** 단계 2 (2-3시간)
