import re

def extract_problematic_buses():
    problem_buses = {}
    current_bus = None
    bus_info = []
    is_reading_bus_info = False
    
    # 파일에서 정보 읽기 (여러 인코딩 시도)
    encodings = ['utf-8', 'cp949', 'euc-kr', 'latin1']
    file_content = None
    
    for encoding in encodings:
        try:
            with open('network_balance.txt', 'r', encoding=encoding) as file:
                file_content = file.readlines()
                break
        except UnicodeDecodeError:
            continue
    
    if file_content is None:
        print("오류: 파일을 읽을 수 없습니다. 인코딩 문제가 있을 수 있습니다.")
        return
    
    # 문제가 있는 버스 정보 분석
    for i, line in enumerate(file_content):
        # 문제가 있는 버스 찾기
        if '[문제]' in line:
            match = re.search(r"\[문제\] 버스 '([^']+)'", line)
            if match:
                current_bus = match.group(1)
                problem_buses[current_bus] = {'reason': line.strip(), 'details': []}
                is_reading_bus_info = True
                continue
        
        # 버스 정보 읽기 중이면 정보 추가
        if is_reading_bus_info and line.strip().startswith('  - '):
            problem_buses[current_bus]['details'].append(line.strip())
        
        # 다음 버스 정보가 시작되면 현재 버스 정보 읽기 종료
        if is_reading_bus_info and line.strip().startswith('['):
            is_reading_bus_info = False
    
    # GND_EL 버스 정보 추출 (BSN_GND 연결 관련)
    gnd_el_info = {}
    is_reading_bus_info = False
    
    for i, line in enumerate(file_content):
        if "버스 'GND_EL'" in line:
            gnd_el_info['reason'] = line.strip()
            gnd_el_info['details'] = []
            is_reading_bus_info = True
            continue
        
        if is_reading_bus_info and line.strip().startswith('  - '):
            gnd_el_info['details'].append(line.strip())
        
        if is_reading_bus_info and line.strip().startswith('['):
            is_reading_bus_info = False
    
    # BSN_EL 버스 정보 추출 (BSN_GND 연결 관련)
    bsn_el_info = {}
    is_reading_bus_info = False
    
    for i, line in enumerate(file_content):
        if "버스 'BSN_EL'" in line:
            bsn_el_info['reason'] = line.strip()
            bsn_el_info['details'] = []
            is_reading_bus_info = True
            continue
        
        if is_reading_bus_info and line.strip().startswith('  - '):
            bsn_el_info['details'].append(line.strip())
        
        if is_reading_bus_info and line.strip().startswith('['):
            is_reading_bus_info = False
    
    # 라인 정보를 분석하여 BSN_GND에 대한 정보 추출
    bsn_gnd_info = []
    in_lines_section = False
    
    for i, line in enumerate(file_content):
        if '라인 정보 확인' in line:
            in_lines_section = True
            continue
        
        if in_lines_section and 'BSN_GND' in line:
            bsn_gnd_info.append(line.strip())
            # 다음 3줄도 저장 (범위 체크)
            for j in range(1, 4):
                if i + j < len(file_content):
                    bsn_gnd_info.append(file_content[i + j].strip())
    
    # 문제가 있는 버스 요약 출력
    print("=" * 50)
    print("문제가 있는 버스 요약")
    print("=" * 50)
    
    problem_summary = False
    for i, line in enumerate(file_content):
        if "문제가 있는 버스 요약" in line:
            problem_summary = True
            print(line.strip())
            # 다음 10줄 출력 (범위 체크)
            for j in range(1, 10):
                if i + j < len(file_content) and file_content[i + j].strip():
                    print(file_content[i + j].strip())
            break
    
    if not problem_summary:
        print("문제가 있는 버스 요약을 찾을 수 없습니다.")
    
    # 결과 출력
    print("\n" + "=" * 50)
    print("문제가 있는 버스 세부 분석")
    print("=" * 50)
    
    if problem_buses:
        for bus, info in problem_buses.items():
            print(f"\n버스: {bus}")
            print(f"문제: {info['reason']}")
            print("\n세부 정보:")
            for detail in info['details']:
                print(detail)
    else:
        print("문제가 있는 버스가 없습니다.")
    
    print("\n" + "=" * 50)
    print("오류 메시지에 언급된 BSN_GND 라인 관련 정보")
    print("=" * 50)
    
    print("\nBSN_EL 버스 정보:")
    if bsn_el_info:
        print(f"상태: {bsn_el_info['reason']}")
        print("\n세부 정보:")
        for detail in bsn_el_info['details']:
            print(detail)
    else:
        print("BSN_EL 버스 정보를 찾을 수 없습니다.")
        
    print("\nGND_EL 버스 정보:")
    if gnd_el_info:
        print(f"상태: {gnd_el_info['reason']}")
        print("\n세부 정보:")
        for detail in gnd_el_info['details']:
            print(detail)
    else:
        print("GND_EL 버스 정보를 찾을 수 없습니다.")
    
    print("\nBSN_GND 라인 정보:")
    if bsn_gnd_info:
        for info in bsn_gnd_info:
            print(info)
    else:
        print("BSN_GND 라인 정보를 찾을 수 없습니다.")

if __name__ == "__main__":
    extract_problematic_buses() 