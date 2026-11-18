"""
2040 대표경로 ESS −30% 조건 MGA 민감도 (ε = 0.5, 1, 1.5, 2%)

기준 결과: results/대표경로_1116/2040_ESS_-30%/optimization_result_20261116_205828.nc
현재 interface.xlsx(입력시트 B36 = −50%)와 기준 결과의 차이는 ESS 저장설비(_ESS, _ESS_RE) 단가뿐이므로,
네트워크 생성 후 해당 단가를 기준 결과 값으로 바꿔(메모리에서만) run_mga_sweep 절차를 그대로 실행한다.

사용법:
  python run_mga_2040_ess30.py [--eps 0.5 1 1.5 2] [--check]
"""
import argparse
import glob
import os
import shutil
import sys
import tempfile

import numpy as np
import pypsa

import run_mga_sweep as r

REF_DIR = os.path.join('results', '대표경로_1116', '2040_ESS_-30%')
ATTRS = {'generators': ['p_nom', 'p_nom_max', 'p_nom_extendable', 'capital_cost', 'marginal_cost', 'efficiency'],
         'links': ['p_nom', 'p_nom_max', 'p_nom_extendable', 'capital_cost', 'marginal_cost', 'efficiency'],
         'stores': ['e_nom', 'e_nom_max', 'e_nom_extendable', 'capital_cost', 'standing_loss'],
         'lines': ['s_nom', 's_nom_extendable', 'capital_cost'],
         'loads': ['bus']}


def load_reference():
    src = glob.glob(os.path.join(REF_DIR, '*.nc'))[0]
    tmp = os.path.join(tempfile.gettempdir(), 'mga2040_ess30_ref.nc')
    shutil.copy2(src, tmp)
    return pypsa.Network(tmp), src


def ess_cost_from(ref):
    s = ref.stores
    idx = [x for x in s.index if x.endswith('_ESS') or x.endswith('_ESS_RE')]
    return s.loc[idx, 'capital_cost']


def compare(n, ref):
    bad = []
    for comp, attrs in ATTRS.items():
        a, b = getattr(n, comp), getattr(ref, comp)
        if set(a.index) != set(b.index):
            bad.append((comp, 'index', len(set(a.index) ^ set(b.index))))
            continue
        for at in attrs:
            x, y = a[at], b.loc[a.index, at]
            if x.dtype == object or y.dtype == object:
                diff = (x.astype(str) != y.astype(str)).sum()
            else:
                diff = (~np.isclose(x.astype(float), y.astype(float), rtol=1e-6, atol=1e-6, equal_nan=True)).sum()
            if diff:
                bad.append((comp, at, int(diff)))
    lp = float(n.loads_t.p_set.sum().sum()), float(ref.loads_t.p_set.sum().sum())
    if abs(lp[0] - lp[1]) > 1e-6 * lp[1]:
        bad.append(('loads_t', 'p_set', lp))
    for comp, tab in (('generators', 'p_max_pu'), ('links', 'p_min_pu'), ('links', 'p_max_pu'), ('stores', 'e_min_pu')):
        A, B = getattr(getattr(n, f'{comp}_t'), tab), getattr(getattr(ref, f'{comp}_t'), tab)
        cols = A.columns.intersection(B.columns)
        if len(cols) and not np.allclose(A[cols].values, B[cols].values, rtol=1e-6, atol=1e-6):
            bad.append((f'{comp}_t', tab, 'values'))
    return bad


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--eps', type=float, nargs='+', default=[0.5, 1, 1.5, 2])
    ap.add_argument('--check', action='store_true', help='네트워크 대조만 하고 종료')
    args = ap.parse_args()

    ref, src = load_reference()
    ess_cost = ess_cost_from(ref)
    print(f"[ESS-30] 기준 결과: {src}\n[ESS-30] 기준 목적함수 {ref.objective / 1e9:.4f}조, "
          f"ESS 단가 {ess_cost.iloc[0]:,.2f} ({len(ess_cost)}개)", flush=True)

    build0 = r.build

    def build(input_data):
        n = build0(input_data)
        n.stores.loc[ess_cost.index, 'capital_cost'] = ess_cost.values
        print(f"[ESS-30] ESS 저장설비 단가 → {ess_cost.iloc[0]:,.2f} ({len(ess_cost)}개)", flush=True)
        return n

    r.build = build
    if args.check:
        n = build(r.load_input())
        bad = compare(n, ref)
        print('[ESS-30] 기준 결과와 차이:', bad or '없음')
        return
    sys.argv = [sys.argv[0], '--eps', *[f'{e:g}' for e in args.eps], '--no-penalty', '--tag', '2040ess30']
    r.main()


if __name__ == '__main__':
    main()
