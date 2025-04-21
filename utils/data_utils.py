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

def preprocess_data(df, license_col=None, name_col=None):
    """데이터 전처리 함수"""
    if license_col is not None and name_col is not None:
        # 선택된 열만 추출
        df = df[[license_col, name_col]].copy()
        # 열 이름 표준화
        df.columns = ['면허번호', '성명']
    
    # NaN 값 제거
    df = df.dropna()
    
    # 면허번호를 문자열로 변환하고 공백 제거
    df['면허번호'] = df['면허번호'].astype(str).str.strip()
    # 성명에서 공백 제거
    df['성명'] = df['성명'].astype(str).str.strip()
    
    return df

def get_data_overview(df1, df2):
    """두 데이터프레임의 개요 정보를 계산하는 함수"""
    # 두 파일 모두에 있는 면허번호 수
    common = pd.merge(df1[['면허번호']], df2[['면허번호']], on='면허번호', how='inner')
    common_count = len(common)
    
    # 첫 번째 파일에만 있는 데이터 (두 번째 파일에는 없는 데이터)
    only_in_first = pd.merge(
        df1, 
        df2[['면허번호']], 
        on='면허번호', 
        how='left', 
        indicator=True
    )
    only_in_first = only_in_first[only_in_first['_merge'] == 'left_only']
    only_in_first_count = len(only_in_first)
    
    # 공통 데이터 계산
    common_data = pd.merge(df1, df2, on='면허번호', how='inner', suffixes=('_1', '_2'))
    
    # 일치 여부 추가
    common_data['상태'] = common_data.apply(
        lambda row: '일치' if row['성명_1'] == row['성명_2'] else '불일치', 
        axis=1
    )
    
    # 일치/불일치 수 계산
    matched_count = len(common_data[common_data['상태'] == '일치'])
    unmatched_count = len(common_data[common_data['상태'] == '불일치'])
    
    # 일치율 계산
    match_rate = (matched_count / common_count) * 100 if common_count > 0 else 0
    
    overview = {
        'first_file_count': len(df1),
        'second_file_count': len(df2),
        'common_count': common_count,
        'only_in_first_count': only_in_first_count,
        'first_minus_common': len(df1) - common_count,
        'matched_count': matched_count,
        'unmatched_count': unmatched_count,
        'match_rate': match_rate,
    }
    
    return overview, common_data, only_in_first