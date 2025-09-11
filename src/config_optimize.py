
# 최적화 설정
solve_opts = {
    "solver_name": "cplex",
    "solver_options": {
        "threads": 4,
        "lpmethod": 4,  # 배리어 메서드
        "barrier.algorithm": 3,  # 기본 배리어 알고리즘
        "mip.tolerances.mipgap": 0.05,  # MIP 갭 허용치
        "timelimit": 3600  # 시간 제한 (초)
    },
    "formulation": "kirchhoff"
}

# 시간 범위 제한 (계산량 감소)
time_settings = {
    "start_time": "2023-01-01 00:00:00",
    "end_time": "2023-01-31 23:00:00",  # 1월 한 달만 사용
    "freq": "1h"
}

# 제약조건 완화
constraints = {}  # 모든 제약조건 비활성화
