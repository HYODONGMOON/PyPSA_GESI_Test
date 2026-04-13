import pypsa

# PyPSA에서 사용 가능한 solver 확인
print("사용 가능한 Solver 확인:")
print("=" * 80)

solvers_to_check = ['cplex', 'gurobi', 'glpk', 'cbc', 'highs']

for solver in solvers_to_check:
    try:
        # 더미 네트워크 생성
        network = pypsa.Network()
        network.set_snapshots(range(2))
        network.add("Bus", "bus")
        network.add("Generator", "gen", bus="bus", p_nom=100, marginal_cost=10)
        network.add("Load", "load", bus="bus", p_set=[50, 50])
        
        # Solver 테스트
        status = network.optimize(solver_name=solver, extra_functionality=None)
        
        if status[0] == 'ok':
            print(f"  [{solver}] 사용 가능 ✓")
        else:
            print(f"  [{solver}] 설치되어 있지만 작동하지 않음: {status}")
            
    except Exception as e:
        error_msg = str(e)
        if 'not installed' in error_msg.lower() or 'not found' in error_msg.lower():
            print(f"  [{solver}] 설치 안됨 ✗")
        else:
            print(f"  [{solver}] 오류: {error_msg}")

print("\n" + "=" * 80)
print("권장 사항:")
print("  1. CPLEX 사용: conda install -c ibmdecisionoptimization cplex")
print("  2. HiGHS 사용 (무료): pip install highspy")
print("  3. GLPK 사용 (무료): conda install -c conda-forge glpk")
