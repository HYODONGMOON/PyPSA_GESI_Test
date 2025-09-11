import sys

def search_lines(file_path, search_terms, encoding='cp949'):
    """파일에서 특정 단어가 포함된 줄을 검색"""
    try:
        with open(file_path, 'r', encoding=encoding) as file:
            lines = file.readlines()
    except UnicodeDecodeError:
        print(f"오류: {encoding} 인코딩으로 파일을 읽을 수 없습니다.")
        return
    
    print(f"'{file_path}' 파일에서 다음 용어 검색: {', '.join(search_terms)}")
    print("=" * 70)
    
    found = False
    for i, line in enumerate(lines):
        # 검색어 중 하나라도 포함된 줄 찾기
        if any(term in line for term in search_terms):
            found = True
            print(f"줄 {i+1}: {line.strip()}")
            # 앞뒤 컨텍스트 5줄 출력
            context_start = max(0, i-5)
            context_end = min(len(lines), i+6)
            
            print("\n컨텍스트:")
            for j in range(context_start, context_end):
                prefix = ">" if j == i else " "
                print(f"{prefix} {j+1}: {lines[j].strip()}")
            print("-" * 70)
    
    if not found:
        print(f"검색어를 포함하는 줄을 찾을 수 없습니다.")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("사용법: python search_lines.py <파일_경로> <검색어1> [검색어2 ...]")
        sys.exit(1)
    
    file_path = sys.argv[1]
    search_terms = sys.argv[2:]
    search_lines(file_path, search_terms) 