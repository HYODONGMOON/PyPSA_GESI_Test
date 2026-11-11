# -*- mode: python ; coding: utf-8 -*-
#
# PyPSA_Analysis.spec
# PyInstaller 빌드 설정 파일
# 사용법: pyinstaller PyPSA_Analysis.spec --clean --noconfirm
#

import os
from PyInstaller.utils.hooks import collect_all, collect_data_files

block_cipher = None

# ── 무거운 패키지 전체 수집 ────────────────────────────────────────────────
pypsa_d,   pypsa_b,   pypsa_h   = collect_all('pypsa')
linopy_d,  linopy_b,  linopy_h  = collect_all('linopy')
highspy_d, highspy_b, highspy_h = collect_all('highspy')
pyproj_d,  pyproj_b,  pyproj_h  = collect_all('pyproj')
geopandas_d, geopandas_b, geopandas_h = collect_all('geopandas')
shapely_d, shapely_b, shapely_h = collect_all('shapely')
fiona_d,   fiona_b,   fiona_h   = collect_all('fiona')
matplotlib_d, matplotlib_b, matplotlib_h = collect_all('matplotlib')
networkx_d, networkx_b, networkx_h = collect_all('networkx')
xarray_d,  xarray_b,  xarray_h  = collect_all('xarray')
scipy_d,   scipy_b,   scipy_h   = collect_all('scipy')

# ── Analysis ──────────────────────────────────────────────────────────────
a = Analysis(
    ['PyPSA_GUI.py'],
    pathex=[
        os.path.abspath('.'),
        os.path.abspath('src'),
    ],
    binaries=(
        pypsa_b + linopy_b + highspy_b + pyproj_b +
        geopandas_b + shapely_b + fiona_b + matplotlib_b +
        networkx_b + xarray_b + scipy_b
    ),
    datas=(
        # ── 소스 모듈 (src/) ────────────────────────────────────────────
        [('src', 'src')] +
        # ── 사용자 입력 파일 ─────────────────────────────────────────────
        [('interface.xlsx', '.')] +
        # integrated_input_data.xlsx 는 interface.xlsx 로부터 생성되므로
        # 존재할 경우에만 포함 (없어도 실행 시 자동 생성됨)
        ([('integrated_input_data.xlsx', '.')] if os.path.exists('integrated_input_data.xlsx') else []) +
        # ── 패키지 데이터 ────────────────────────────────────────────────
        pypsa_d + linopy_d + highspy_d + pyproj_d +
        geopandas_d + shapely_d + fiona_d + matplotlib_d +
        networkx_d + xarray_d + scipy_d
    ),
    hiddenimports=(
        # 핵심 과학계산
        ['pypsa', 'linopy', 'highspy', 'xarray', 'xarray.backends',
         'scipy', 'scipy.optimize', 'scipy.sparse', 'scipy.linalg',
         'numpy', 'pandas', 'openpyxl', 'openpyxl.styles',
         'openpyxl.utils', 'openpyxl.drawing',
         # 지리
         'pyproj', 'pyproj.transformer', 'pyproj.datadir',
         'geopandas', 'shapely', 'shapely.geometry', 'fiona',
         'pyogrio',
         # 시각화
         'matplotlib', 'matplotlib.pyplot', 'matplotlib.backends',
         'matplotlib.backends.backend_agg',
         # 기타
         'networkx', 'copy', 'traceback',
         # src 모듈
         'korea_map', 'process_regional_excel', 'regional_data_manager',
         'PyPSA_HD_Model', 'PyPSA_HD_Regional',
        ] +
        pypsa_h + linopy_h + highspy_h + pyproj_h +
        geopandas_h + shapely_h + fiona_h + matplotlib_h +
        networkx_h + xarray_h + scipy_h
    ),
    hookspath=[],
    hooksconfig={},
    # cplex mock 훅: Python 시작 직후 cplex 자리에 mock 삽입
    # → linopy/solvers.py 의 'import cplex' 가 DLL 오류 없이 통과
    runtime_hooks=['rthooks/rthook_mock_cplex.py'],
    excludes=[
        # cplex: 핵심 DLL 이 conda env 외부에 있어 번들 불가
        #        runtime hook 이 mock 으로 대체하므로 제외해도 무방
        'cplex',
        # 불필요한 패키지 제외 (용량 절감)
        'tkinter', 'PyQt5', 'PyQt6', 'PySide2', 'PySide6',
        'IPython', 'jupyter', 'notebook', 'pytest',
        'sphinx', 'docutils',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PyPSA_Analysis',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,          # 콘솔 창 표시 (분석 진행 로그 확인용)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='PyPSA_Analysis',
)
