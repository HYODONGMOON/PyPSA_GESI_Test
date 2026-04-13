# CPLEX 사용 가이드

## CPLEX가 인식되지 않는 이유

CPLEX가 인식되지 않는 주된 이유는 **잘못된 Python 환경**에서 실행했기 때문입니다:

### 문제 상황
```
Python 환경: C:\Program Files\Python313\python.exe (시스템 Python)
사용 가능한 솔버: highs, cbc (CPLEX 없음)
```

### 원인
- CPLEX는 `pypsa_env` conda 환경에만 설치되어 있습니다
- 시스템 Python으로 실행하면 CPLEX를 찾을 수 없습니다

## 솔버 성능 비교

### 대규모 문제 (8760시간, 500개 버스, 3500개 선로)

| 솔버 | 예상 실행 시간 | 메모리 사용량 | 최적화 품질 |
|------|---------------|--------------|-------------|
| **CPLEX** | **2-4시간** | 30-40GB | ⭐⭐⭐⭐⭐ 최고 |
| Gurobi | 2-5시간 | 30-40GB | ⭐⭐⭐⭐⭐ 최고 |
| HiGHS | 20-40시간 | 20-30GB | ⭐⭐⭐⭐ 우수 |
| CBC | 40-80시간 | 25-35GB | ⭐⭐⭐ 보통 |
| GLPK | 80시간+ | 30GB+ | ⭐⭐ 낮음 |

**결론**: CPLEX는 다른 무료 솔버보다 **10-20배 빠릅니다**. 대규모 문제에서는 필수입니다.

## 해결 방법

### 방법 1: 자동 재실행 (권장) ✅

이제 스크립트가 자동으로 CPLEX가 설치된 conda 환경으로 재실행됩니다!

**아무 Python 환경에서나 실행하세요:**
```powershell
python PyPSA_GUI.py --rolling-horizon
```

스크립트가 자동으로:
1. CPLEX 가용성 확인
2. conda 환경 경로 탐색
3. conda 환경의 Python으로 재실행

### 방법 2: 배치 파일 사용 (가장 간단) 🚀

**전체 분석 (8760시간):**
```
run_full_analysis.bat 더블클릭
```

**Rolling Horizon (12개 구간):**
```
run_rolling_horizon.bat 더블클릭
```

**Rolling Horizon (4개 구간):**
```
run_rolling_horizon_4segments.bat 더블클릭
```

### 방법 3: conda 환경에서 직접 실행

**Anaconda Prompt 열기:**
```powershell
conda activate pypsa_env
cd "C:\Users\Hyodong.Moon\Desktop\HDMOON\python workplace\PyPSA_GESI_Test"

# 전체 분석
python PyPSA_GUI.py

# Rolling Horizon (12개 구간)
python PyPSA_GUI.py --rolling-horizon

# Rolling Horizon (4개 구간)
python PyPSA_GUI.py --rolling-horizon --segments 4
```

### 방법 4: PowerShell에서 직접 실행

```powershell
# conda 환경의 Python 직접 호출
C:\ProgramData\anaconda3\envs\pypsa_env\python.exe PyPSA_GUI.py --rolling-horizon
```

## CPLEX 설치 확인

conda 환경에서 CPLEX가 제대로 설치되어 있는지 확인:

```powershell
conda activate pypsa_env
python -c "from linopy import solvers; print(solvers.available_solvers)"
```

**예상 출력:**
```
['cplex', 'gurobi', 'highs', 'cbc', 'glpk']
```

`cplex`가 목록에 있으면 정상입니다.

## 문제 해결

### CPLEX가 여전히 인식되지 않는 경우

1. **conda 환경 확인:**
   ```powershell
   conda env list
   ```
   `pypsa_env`가 목록에 있는지 확인

2. **CPLEX 재설치:**
   ```powershell
   conda activate pypsa_env
   conda install -c ibmdecisionoptimization cplex
   ```

3. **환경 변수 확인:**
   - CPLEX 라이선스 파일 위치 확인
   - `ILOG_LICENSE_FILE` 환경 변수 설정

## 권장 실행 방식

### 대규모 문제 (8760시간 전체)
- ❌ **비권장**: 전체 분석 (메모리 부족 가능)
- ✅ **권장**: Rolling Horizon 4-12개 구간

### Rolling Horizon 구간 수 선택 가이드

| 구간 수 | 구간당 시간 | 메모리 사용 | 실행 시간 | 권장 용도 |
|---------|------------|------------|----------|----------|
| 4개 | 2190시간 | ~15GB | 중간 | ⭐ 빠른 분석 |
| 6개 | 1460시간 | ~10GB | 중간 | ⭐⭐ 균형잡힌 |
| 12개 | 730시간 | ~5GB | 긴 | ⭐⭐⭐ 안전한 |
| 24개 | 365시간 | ~3GB | 매우 긴 | 메모리 부족 시 |

**권장**: 4-6개 구간 (CPLEX 사용 시 충분히 빠름)

## 성능 최적화 팁

1. **CPLEX 사용** (필수)
2. **적절한 구간 수 선택** (4-6개 권장)
3. **충분한 메모리 확보** (최소 32GB RAM)
4. **SSD 사용** (I/O 성능 향상)
5. **백그라운드 프로그램 종료** (메모리 확보)

## 추가 도움말

문제가 계속되면:
1. 터미널 로그 확인
2. Python 환경 경로 확인
3. CPLEX 라이선스 상태 확인
