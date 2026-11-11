"""
Runtime hook: exe 실행 시 cplex를 가짜(mock)로 대체.

배경:
  linopy/solvers.py 가 모듈 로드 시점에 'import cplex' 를 실행함.
  exe 번들에는 CPLEX 핵심 DLL(py310_cplex2211)이 없어서 즉시 충돌 발생.
  이 훅이 Python 시작 직후 cplex 자리에 mock 을 넣어두면
  linopy 가 정상적으로 로드되고, 이후 _detect_solver() 가 mock 여부를
  확인해서 HiGHS 로 자동 전환함.
"""
import sys


class _CplexMock:
    """linopy 의 cplex import 를 통과시키는 최소 mock."""

    _IS_MOCK_FOR_EXE = True  # _detect_solver() 가 이 속성으로 mock 여부 판별

    def __init__(self, *args, **kwargs):
        pass

    def __getattr__(self, name):
        return _CplexMock()

    def __call__(self, *args, **kwargs):
        return _CplexMock()

    def __iter__(self):
        return iter([])

    def __bool__(self):
        return False

    def __repr__(self):
        return "<CplexMock>"


_mock = _CplexMock()

# linopy/solvers.py 가 참조하는 cplex 서브모듈 전부 mock 으로 등록
_cplex_modules = [
    'cplex',
    'cplex.aborter',
    'cplex.callbacks',
    'cplex.exceptions',
    'cplex.six',
    'cplex.six.moves',
    'cplex._internal',
    'cplex._internal._list_array_utils',
    'cplex._internal._pycplex',
    'cplex._internal._pycplex_platform',
]
for _mod in _cplex_modules:
    sys.modules[_mod] = _mock
