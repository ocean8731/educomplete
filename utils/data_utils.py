import pandas as pd
from collections import defaultdict

def transform_and_split(df):
    """
    데이터프레임을 변환하고 분할하는 함수
    
    Args:
        df: 변환할 데이터프레임
    
    Returns:
        list: 변환된 데이터프레임 리스트
    
    Raises:
        ValueError: 필요한 컬럼이 없는 경우
    """
    required = ['면허번호', '성명', '이수년도']
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"❌ 필요한 컬럼이 누락되었습니다: {missing}")
    
    df = df[required].copy()
    df['이수년도'] = df['이수년도'].astype(str)
    grouped = df.groupby(['면허번호', '성명'])

    index_map = defaultdict(int)
    file_data = []

    for (lic, name), group in grouped:
        group = group.sort_values(by='이수년도')
        for _, row in group.iterrows():
            idx = index_map[(lic, name)]
            if len(file_data) <= idx:
                file_data.append([])
            file_data[idx].append(row)
            index_map[(lic, name)] += 1

    output_dfs = []
    for file_rows in file_data:
        out_df = pd.DataFrame(file_rows).reset_index(drop=True)
        out_df.insert(0, "순번", range(1, len(out_df) + 1))
        output_dfs.append(out_df)

    return output_dfs

def improved_filter_data(df, filter_year):
    """
    데이터를 필터링하는 개선된 함수
    
    Args:
        df: 원본 데이터프레임
        filter_year: 필터링할 연도 또는 "모든 연도"
        
    Returns:
        tuple: (필터링된 데이터프레임, 필터링 메시지)
    """
    # 원본 데이터 정보
    print(f"원본 데이터: {len(df)}행 x {len(df.columns)}열")
    
    # 빈 값 처리
    df = df.replace('', pd.NA)
    
    # 필터링 전 데이터 상태 확인
    if '이수년도' not in df.columns:
        return df, "❌ '이수년도' 열이 없습니다!"
    
    # 이수년도 열 로깅
    unique_years = df['이수년도'].unique()
    print(f"이수년도 고유값: {unique_years.tolist()}")
    
    # 각 이수년도별 행 수 세기
    year_counts = df['이수년도'].value_counts().to_dict()
    print("이수년도별 행 수:")
    for year, count in year_counts.items():
        print(f"  - {year}: {count}행")
    
    # 필터링 적용
    if filter_year != "모든 연도":
        # 연도 타입 변환 전 확인
        print(f"필터링할 연도: {filter_year} (타입: {type(filter_year)})")
        print(f"데이터의 이수년도 타입: {df['이수년도'].dtype}")
        
        # 문자열로 통일하여 비교
        df['이수년도_str'] = df['이수년도'].astype(str)
        filter_year_str = str(filter_year)
        
        # 필터링 조건 로깅
        filter_condition = df['이수년도_str'] != filter_year_str
        print(f"필터 조건 결과: {filter_condition.sum()}/{len(df)}행 선택됨")
        
        # 필터링 적용
        df_filtered = df[filter_condition].copy()
        df_filtered = df_filtered.drop(columns=['이수년도_str'])
        
        filter_message = f"{filter_year}년 데이터 제외 필터링"
    else:
        df_filtered = df.copy()
        filter_message = "모든 연도 데이터 포함 (필터링 없음)"
    
    # 필터링 결과 로깅
    print(f"필터링 결과: {len(df_filtered)}행 남음 (제외된 행: {len(df) - len(df_filtered)})")
    
    # 빈 행 제거
    df_filtered = df_filtered.dropna(how='all')
    print(f"빈 행 제거 후: {len(df_filtered)}행")
    
    return df_filtered, filter_message