
# ����ȭ ����
solve_opts = {
    "solver_name": "cplex",
    "solver_options": {
        "threads": 4,
        "lpmethod": 4,  # �踮�� �޼���
        "barrier.algorithm": 3,  # �⺻ �踮�� �˰���
        "mip.tolerances.mipgap": 0.05,  # MIP �� ���ġ
        "timelimit": 3600  # �ð� ���� (��)
    },
    "formulation": "kirchhoff"
}

# �ð� ���� ���� (��귮 ����)
time_settings = {
    "start_time": "2023-01-01 00:00:00",
    "end_time": "2023-01-31 23:00:00",  # 1�� �� �޸� ���
    "freq": "1h"
}

# �������� ��ȭ
constraints = {}  # ��� �������� ��Ȱ��ȭ
