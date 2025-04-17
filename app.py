import streamlit as st
import pandas as pd
import io
import zipfile
import xlwt
import holidays
from datetime import datetime, date, timedelta
from bs4 import BeautifulSoup
import traceback

st.set_page_config(page_title="엑셀 자동화 도구", layout="wide")
st.title("📄 엑셀 자동화 도구")

# 날짜 형식 함수
def get_last_workday_str():
    kr_holidays = holidays.KR(years=date.today().year)
    day = date.today() - timedelta(days=1)
    while day in kr_holidays or day.weekday() >= 5:
        day -= timedelta(days=1)

    weekdays_kr = ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일']
    day_str = day.strftime("%y.%m.%d.")
    weekday_str = weekdays_kr[day.weekday()]
    return f"{day_str} {weekday_str}"

today_str = get_last_workday_str()

# 파일명 생성 함수
def get_output_filename(index=None):
    base = f"({today_str}_대상) 현장 보수교육 대체 사이버 교육 이수자 리스트"
    return f"{base}.xlsx" if index is None else f"{base}_{index}.xlsx"

# HTML .xls 파싱
def decode_bytes(file_bytes):
    for encoding in ['utf-8', 'euc-kr', 'cp949']:
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("❌ 파일 인코딩을 확인할 수 없습니다.")

def parse_html_xls(file_bytes, use_first_row_as_header):
    html = decode_bytes(file_bytes)
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('table')
    if table is None:
        raise ValueError("HTML 테이블을 찾을 수 없습니다.")
    df = pd.read_html(str(table), header=None)[0]
    if use_first_row_as_header:
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
    else:
        df.columns = [f"컬럼{i}" for i in range(df.shape[1])]
    df = df.fillna('')
    return df

# 변환 및 분할
def transform_and_split(df):
    required = ['면허번호', '성명', '이수년도']
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"❌ 필요한 컬럼이 누락되었습니다: {missing}")
    
    df = df[required].copy()
    df['이수년도'] = df['이수년도'].astype(str)
    grouped = df.groupby(['면허번호', '성명'])

    from collections import defaultdict
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

def df_to_xlsx_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    output.seek(0)
    return output

# 엑셀 매핑 기능
def convert_excel_format(input_file):
    try:
        df = pd.read_excel(input_file)
        header_mapping = {'처리일자(판정일자)': '처리일자'}
        df_renamed = df.rename(columns=header_mapping)

        output_columns = [
            '보수교육 이수 대상여부', '신청분류', '직종', '이름', '면허(자격)번호', '신청대상연도',
            '신청일자', '면허검증결과', '확인결과', '처리일자', '면제·유예사유', '대리신고 여부', '연락처'
        ]
        available_columns = [col for col in output_columns if col in df_renamed.columns]

        confirmation_result = "unknown"
        if '확인결과' in df_renamed.columns and len(df_renamed) > 0:
            result = df_renamed['확인결과'].iloc[0]
            confirmation_result = str(result) if pd.notna(result) else "unknown"

        df_final = df_renamed[available_columns]
        return df_final, confirmation_result
    except Exception as e:
        st.error(f"변환 중 오류 발생: {str(e)}")
        return None, "error"
# 개선된 변환 함수
def improved_transform_to_new_format(df_filtered):
    """개선된 변환 함수 - 버전 2"""
    
    # 원본 데이터 디버깅
    st.write("원본 데이터 샘플:", df_filtered.head())
    
    # 빈 문자열을 NA로 변환
    df_filtered = df_filtered.replace('', pd.NA)
    
    # 필수 열이 모두 있는지 확인
    required_cols = ['면허번호', '성명', '이수년도']
    if not all(col in df_filtered.columns for col in required_cols):
        missing = [col for col in required_cols if col not in df_filtered.columns]
        st.error(f"필수 열이 누락됨: {missing}")
        return pd.DataFrame()
    
    # 데이터 직접 할당 (대량 변환 방식)
    data = []
    for idx, row in df_filtered.iterrows():
        # 필수 값 확인
        if pd.isna(row['면허번호']) or pd.isna(row['성명']) or pd.isna(row['이수년도']):
            st.write(f"행 {idx} 제외: 필수 값 누락 - {row['성명']} {row['면허번호']}")
            continue
            
        # 새 행 구성 - 딕셔너리로 생성
        new_row = {
            '연번': len(data) + 1,
            '성명': row['성명'],
            '면허번호': row['면허번호'],
            '직종': 15,
            '보수교육 이수년도': row['이수년도'],
            '보수교육 이수시간': row['총평점'] if '총평점' in row and not pd.isna(row['총평점']) else 0,
            '장기 휴직자구분': 1,  # 기본값
            '필수교육 이수시간': 1 if pd.to_numeric(row['이수년도'], errors='coerce') > 2019 else 0,
            '회원 아이디': ''
        }
        
        # 장기 휴직자구분 매핑 (있는 경우)
        if '대상구분' in row and not pd.isna(row['대상구분']):
            mapping = {'대상': 1, '비대상1': 2, '비대상2': 3, '비대상3': 4}
            new_row['장기 휴직자구분'] = mapping.get(row['대상구분'], 1)
        
        # 데이터 추가
        data.append(new_row)
    
    # 데이터가 없으면 빈 데이터프레임 반환
    if not data:
        st.error("변환할 유효한 데이터가 없습니다.")
        return pd.DataFrame()
    
    # 데이터프레임 생성
    new_df = pd.DataFrame(data)
    
    # 결과 출력
    st.write(f"변환된 데이터: {len(new_df)}행")
    st.write("변환 결과 샘플:", new_df.head())
    
    return new_df
    """
    필터링된 데이터를 새 포맷으로 변환하는 개선된 함수
    
    Args:
        df_filtered: 필터링된 데이터프레임
        
    Returns:
        새 포맷으로 변환된 데이터프레임
    """
    # 디버깅을 위한 정보 표시
    st.write("▶️ 변환 전 데이터 상세 정보:")
    st.write(f"- 행 수: {len(df_filtered)}")
    st.write(f"- 컬럼: {df_filtered.columns.tolist()}")
    
    # 각 열의 값 목록 출력 (디버깅용)
    for col in df_filtered.columns:
        unique_values = df_filtered[col].unique()
        if len(unique_values) < 10:  # 고유값이 너무 많지 않은 경우에만 표시
            st.write(f"- '{col}' 열 고유값: {unique_values.tolist()}")
    
    # 데이터 프레임의 처음 몇 개 행을 표시 (디버깅용)
    st.write("원본 데이터 샘플:")
    st.dataframe(df_filtered.head())
    
    # 빈 행이나 모든 열이 빈 문자열인 행 제거 (중요!)
    df_filtered = df_filtered.replace('', pd.NA)
    df_filtered = df_filtered.dropna(how='all')
    
    # 모든 열에 대해 문자열로 변환 후 공백 제거
    for col in df_filtered.columns:
        if df_filtered[col].dtype == 'object':
            df_filtered[col] = df_filtered[col].astype(str).str.strip()
            # 빈 문자열을 NA로 변환
            df_filtered.loc[df_filtered[col] == '', col] = pd.NA
    
    # NA 값 체크 및 처리
    na_counts = df_filtered.isna().sum()
    if na_counts.sum() > 0:
        st.warning("⚠️ 다음 열에 NA 값이 있습니다. 자동으로 처리됩니다:")
        st.write(na_counts[na_counts > 0])
    
    # 필수 컬럼 확인 - NA가 없는지 검증
    required_cols = ['면허번호', '성명', '이수년도']
    for col in required_cols:
        if col in df_filtered.columns and df_filtered[col].isna().any():
            # NA 값이 있는 행 찾기
            na_rows = df_filtered[df_filtered[col].isna()]
            st.warning(f"⚠️ '{col}' 열에 {len(na_rows)}개의 NA 값이 있습니다. 이 행들은 제외됩니다.")
            # NA 값이 있는 행 삭제
            df_filtered = df_filtered.dropna(subset=[col])
    
    # 빈 데이터프레임 체크
    if len(df_filtered) == 0:
        st.error("❌ 유효한 데이터가 없습니다. 모든 행이 필터링되었거나 필수 데이터가 없습니다.")
        return pd.DataFrame()
    
    # 새로운 데이터프레임 생성 - 필요한 행만 포함
    valid_indices = []
    invalid_rows = []
    
    # 각 행 검증 및 로깅
    for idx, row in df_filtered.iterrows():
        # 이 행이 유효한지 확인
        if pd.isna(row['면허번호']) or pd.isna(row['성명']) or pd.isna(row['이수년도']):
            invalid_rows.append({
                '인덱스': idx,
                '면허번호': row['면허번호'] if '면허번호' in row else None,
                '성명': row['성명'] if '성명' in row else None,
                '이수년도': row['이수년도'] if '이수년도' in row else None,
                '이유': '필수 정보 누락'
            })
        else:
            valid_indices.append(idx)
    
    # 유효하지 않은 행 정보 표시 (있는 경우)
    if invalid_rows:
        st.warning(f"⚠️ {len(invalid_rows)}개 행이 유효하지 않아 제외됩니다.")
        st.dataframe(pd.DataFrame(invalid_rows))
    
    # 유효한 행만 선택
    df_valid = df_filtered.loc[valid_indices].copy()
    
    # 디버깅 정보 - 유효한 데이터 상태
    st.write(f"▶️ 검증 후 유효한 데이터: {len(df_valid)}행")
    
    # 새로운 데이터프레임 생성
    new_df = pd.DataFrame()
    
    # 연번 자동 생성
    new_df['연번'] = range(1, len(df_valid) + 1)
    
    # 데이터 매핑 - 명시적으로 각 컬럼 로깅
    st.write("▶️ 데이터 매핑 시작")
    
    # 성명 열 매핑
    st.write(f"  - 성명 매핑: {df_valid['성명'].nunique()}명")
    new_df['성명'] = df_valid['성명'].fillna("미상")
    
    # 면허번호 열 매핑
    st.write(f"  - 면허번호 매핑: {df_valid['면허번호'].nunique()}개")
    new_df['면허번호'] = df_valid['면허번호'].fillna("00000")
    
    # 직종 열 매핑 (고정값 15)
    new_df['직종'] = 15
    
    # 이수년도 열 매핑
    st.write(f"  - 이수년도 매핑: {df_valid['이수년도'].unique().tolist()}")
    new_df['보수교육 이수년도'] = df_valid['이수년도'].astype(str).fillna("0000")
    
    # 총평점 열 처리
    if '총평점' in df_valid.columns:
        try:
            # 모든 값을 로깅
            st.write(f"  - 총평점 값: {df_valid['총평점'].unique().tolist()}")
            new_df['보수교육 이수시간'] = pd.to_numeric(df_valid['총평점'], errors='coerce').fillna(0)
        except Exception as e:
            st.warning(f"총평점 변환 중 오류: {e}. 기본값 0으로 설정합니다.")
            new_df['보수교육 이수시간'] = 0
    else:
        st.warning("'총평점' 열이 없어 보수교육 이수시간을 0으로 설정합니다.")
        new_df['보수교육 이수시간'] = 0
    
    # 장기 휴직자구분 매핑 - 안전하게 처리
    if '대상구분' in df_valid.columns:
        # 매핑 딕셔너리 사용
        mapping = {
            '대상': 1,
            '비대상1': 2,
            '비대상2': 3,
            '비대상3': 4
        }
        
        # 고유값 로깅
        st.write(f"  - 대상구분 고유값: {df_valid['대상구분'].unique().tolist()}")
        
        # 매핑되지 않은 값 처리
        unmapped = [val for val in df_valid['대상구분'].unique() if val not in mapping and not pd.isna(val)]
        if unmapped:
            st.warning(f"⚠️ 알 수 없는 대상구분 값: {unmapped}. 기본값 1로 설정합니다.")
            for val in unmapped:
                mapping[val] = 1
        
        # 안전한 매핑 - 기본값은 1
        new_df['장기 휴직자구분'] = df_valid['대상구분'].map(mapping).fillna(1)
    else:
        st.warning("'대상구분' 열이 없어 장기 휴직자구분을 기본값 1로 설정합니다.")
        new_df['장기 휴직자구분'] = 1
    
    # 필수교육 이수시간 설정 - 안전하게 처리
    try:
        # 연도 데이터가 정수로 변환 가능한지 확인
        years = pd.to_numeric(df_valid['이수년도'], errors='coerce')
        st.write(f"  - 이수년도 변환 결과: {years.unique().tolist()}")
        
        # NA 값은 2019로 처리 (필수교육 이수시간 = 0)
        years = years.fillna(2019) 
        
        # 2019년 이하면 0, 이후면 1
        new_df['필수교육 이수시간'] = (years > 2019).astype(int)
        st.write(f"  - 필수교육 이수시간 분포: {new_df['필수교육 이수시간'].value_counts().to_dict()}")
    except Exception as e:
        st.warning(f"이수년도 변환 중 오류: {e}. 기본값 1로 설정합니다.")
        new_df['필수교육 이수시간'] = 1
    
    # 회원 아이디 칼럼 추가 (비워둠)
    new_df['회원 아이디'] = ''
    
    # 각 열을 안전하게 정수형으로 변환
    try:
        new_df['연번'] = pd.to_numeric(new_df['연번'], errors='coerce').fillna(0).astype(int)
    except Exception as e:
        st.warning(f"연번 변환 오류: {e}")
        
    try:
        new_df['직종'] = pd.to_numeric(new_df['직종'], errors='coerce').fillna(15).astype(int)
    except Exception as e:
        st.warning(f"직종 변환 오류: {e}")
        
    try:
        new_df['장기 휴직자구분'] = pd.to_numeric(new_df['장기 휴직자구분'], errors='coerce').fillna(1).astype(int)
    except Exception as e:
        st.warning(f"장기 휴직자구분 변환 오류: {e}")
        
    try:
        new_df['필수교육 이수시간'] = pd.to_numeric(new_df['필수교육 이수시간'], errors='coerce').fillna(1).astype(int)
    except Exception as e:
        st.warning(f"필수교육 이수시간 변환 오류: {e}")
    
    # 최종 변환 결과 요약
    st.write("▶️ 변환 후 데이터:")
    st.write(f"- 행 수: {len(new_df)}")
    st.write(f"- 면허번호 유니크 값 수: {new_df['면허번호'].nunique()}")
    
    # 변환 결과 미리보기
    st.write("최종 변환 결과 샘플:")
    st.dataframe(new_df.head())
    
    return new_df

# 개선된 필터링 로직
def improved_filter_data(df, filter_year):
    """
    데이터를 필터링하는 개선된 함수
    
    Args:
        df: 원본 데이터프레임
        filter_year: 필터링할 연도 또는 "모든 연도"
        
    Returns:
        필터링된 데이터프레임
    """
    # 원본 데이터 정보
    st.write(f"원본 데이터: {len(df)}행 x {len(df.columns)}열")
    
    # 빈 값 처리
    df = df.replace('', pd.NA)
    
    # 필터링 전 데이터 상태 확인
    if '이수년도' not in df.columns:
        st.error("❌ '이수년도' 열이 없습니다!")
        return df  # 원본 반환
    
    # 이수년도 열 로깅
    unique_years = df['이수년도'].unique()
    st.write(f"이수년도 고유값: {unique_years.tolist()}")
    
    # 각 이수년도별 행 수 세기
    year_counts = df['이수년도'].value_counts().to_dict()
    st.write("이수년도별 행 수:")
    for year, count in year_counts.items():
        st.write(f"  - {year}: {count}행")
    
    # 필터링 적용
    if filter_year != "모든 연도":
        # 연도 타입 변환 전 확인
        st.write(f"필터링할 연도: {filter_year} (타입: {type(filter_year)})")
        st.write(f"데이터의 이수년도 타입: {df['이수년도'].dtype}")
        
        # 문자열로 통일하여 비교
        df['이수년도_str'] = df['이수년도'].astype(str)
        filter_year_str = str(filter_year)
        
        # 필터링 조건 로깅
        filter_condition = df['이수년도_str'] != filter_year_str
        st.write(f"필터 조건 결과: {filter_condition.sum()}/{len(df)}행 선택됨")
        
        # 필터링 적용
        df_filtered = df[filter_condition].copy()
        df_filtered = df_filtered.drop(columns=['이수년도_str'])
        
        filter_message = f"{filter_year}년 데이터 제외 필터링"
    else:
        df_filtered = df.copy()
        filter_message = "모든 연도 데이터 포함 (필터링 없음)"
    
    # 필터링 결과 로깅
    st.write(f"필터링 결과: {len(df_filtered)}행 남음 (제외된 행: {len(df) - len(df_filtered)})")
    
    # 빈 행 제거
    df_filtered = df_filtered.dropna(how='all')
    st.write(f"빈 행 제거 후: {len(df_filtered)}행")
    
    return df_filtered, filter_message
# 탭 구성
tab1, tab2, tab3 = st.tabs(["💡 현장대체온라인교육", "📊 면제/유예/비대상", "🔄 보수교육 완료자 명단등록"])

# 탭 1: 현장대체온라인교육
with tab1:
    st.markdown("### 현장대체온라인교육")
    uploaded_html = st.file_uploader("📤 현장대체온라인교육 .xls 파일 업로드", type=["xls"], key="html")

    if uploaded_html:
        use_first_row_as_header = st.checkbox("🔠 첫 번째 행을 컬럼명으로 사용할까요?", value=True)
        try:
            df = parse_html_xls(uploaded_html.read(), use_first_row_as_header)
            if '소속' in df.columns:
                df = df.drop(columns=['소속'])

            st.dataframe(df.astype(str))
            output_dfs = transform_and_split(df)

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for i, df_part in enumerate(output_dfs):
                    xlsx_bytes = df_to_xlsx_bytes(df_part)
                    filename = get_output_filename(None if i == 0 else i + 1)
                    zip_file.writestr(filename, xlsx_bytes.read())

            zip_buffer.seek(0)
            st.download_button(
                label="📦 ZIP 파일 다운로드",
                data=zip_buffer,
                file_name=f"({today_str}) 현장 보수교육 대체 사이버 교육 이수자 리스트.zip",
                mime="application/zip"
            )
        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")

# 탭 2: 면제/유예/비대상
with tab2:
    st.markdown("### 2. 면제/유예/비대상")
    uploaded_excel = st.file_uploader("📥 면제/유예/비대상 .xlsx 또는 .xls 파일 업로드", type=["xlsx", "xls"], key="excel")

    if uploaded_excel:
        df_preview = pd.read_excel(uploaded_excel)
        st.dataframe(df_preview.head())

        if st.button("변환하기"):
            uploaded_excel.seek(0)
            converted_df, confirmation_result = convert_excel_format(uploaded_excel)

            if converted_df is not None:
                st.subheader("🔍 변환된 데이터")
                st.dataframe(converted_df.head(10))

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    converted_df.to_excel(writer, index=False, sheet_name='Data')
                buffer.seek(0)

                today = datetime.now().strftime("%m%d")
                download_filename = f"Data {today} {confirmation_result}.xlsx"

                st.download_button(
                    label="📥 변환된 파일 다운로드",
                    data=buffer,
                    file_name=download_filename,
                    mime="application/vnd.ms-excel"
                )

                st.success(f"총 {len(converted_df)}개 행 변환 완료!")
                st.info(f"파일명: {download_filename}")
                if '이름' in converted_df.columns:
                    st.write(f"- 고유 사용자 수: {converted_df['이름'].nunique()}")

# 탭 3: 이수년도 필터링 (개선된 버전)
with tab3:
    st.markdown("### 3. 보수교육 완료자 명단등록")
    
    # 상태 관리용 변수
    if 'df_original' not in st.session_state:
        st.session_state.df_original = None
    if 'df_filtered' not in st.session_state:
        st.session_state.df_filtered = None
    if 'df_transformed' not in st.session_state:
        st.session_state.df_transformed = None
    
    # 파일 업로드
    uploaded_html_filter = st.file_uploader("📤 필터링할 .xls HTML 파일 업로드", type=["xls"], key="html_filter")
    
    # 필터링 설정
    col1, col2 = st.columns(2)
    with col1:
        use_first_row = st.checkbox("🔠 첫 번째 행을 컬럼명으로 사용", value=True, key="filter_checkbox")
    with col2:
        current_year = datetime.now().year
        filter_year = st.selectbox("필터링할 연도", 
                                  options=[current_year, current_year-1, current_year-2, "모든 연도"],
                                  index=0,
                                  help="선택한 연도의 데이터를 제외합니다. '모든 연도'를 선택하면 필터링하지 않습니다.")
    
    # 고급 옵션 표시 여부
    show_advanced = st.checkbox("고급 옵션 표시", value=False)
    
    if show_advanced:
        # 고급 데이터 검증 옵션
        st.subheader("🔧 고급 데이터 검증 옵션")
        col1, col2 = st.columns(2)
        with col1:
            remove_empty_rows = st.checkbox("빈 행 자동 제거", value=True, 
                                           help="모든 값이 비어있거나 NA인 행을 자동으로 제거합니다.")
        with col2:
            validate_required_fields = st.checkbox("필수 필드 검증", value=True,
                                                help="면허번호, 성명, 이수년도 필드가 비어있는 행을 제외합니다.")
    else:
        # 기본값 설정
        remove_empty_rows = True
        validate_required_fields = True
    
    if uploaded_html_filter:
        try:
            # 진행 상태 표시
            with st.spinner("파일 파싱 중..."):
                # 파일 유효성 검사
                file_content = uploaded_html_filter.read()
                if len(file_content) == 0:
                    st.error("❌ 빈 파일입니다.")
                else:
                    # HTML 파일 파싱
                    df = parse_html_xls(file_content, use_first_row)
                    st.session_state.df_original = df.copy()
                    
                    # 데이터 구조 검증
                    st.success("✅ 파일 파싱 완료")
                    
                    # 원본 데이터 컬럼 정보 표시
                    st.subheader("🔍 원본 데이터 컬럼 정보")
                    col_info = pd.DataFrame({
                        "컬럼명": df.columns,
                        "데이터 타입": df.dtypes.astype(str),
                        "샘플값": [str(df[col].iloc[0]) if len(df) > 0 else "" for col in df.columns],
                        "고유값 수": [df[col].nunique() for col in df.columns],
                        "결측값 수": [df[col].isna().sum() for col in df.columns]
                    })
                    st.dataframe(col_info)
                    
                    # 원본 데이터 미리보기
                    st.subheader("🔍 업로드된 데이터 미리보기")
                    st.dataframe(df, height=300)
                    
                    # 필요한 열이 있는지 확인
                    required_columns = ['면허번호', '성명', '이수년도']
                    missing = [col for col in required_columns if col not in df.columns]
                    
                    if missing:
                        st.error(f"❌ 필요한 컬럼이 누락되었습니다: {missing}")
                        st.info("업로드된 파일의 컬럼명을 확인하세요. 첫 번째 행을 컬럼명으로 사용하는 옵션을 조정해보세요.")
                    else:
                        # 필터링 및 변환 버튼
                        if st.button("필터링 및 변환 실행"):
                            # 진행 상태 표시 컨테이너
                            progress_container = st.container()
                            with progress_container:
                                progress_bar = st.progress(0)
                                status_text = st.empty()
                            
                            # 로그 표시용 컨테이너
                            log_container = st.expander("상세 로그 보기", expanded=False)
                            
                            # 1단계: 데이터 필터링
                            status_text.text("1/3 단계: 데이터 필터링 중...")
                            
                            with log_container:
                                st.subheader("📋 필터링 과정 로그")
                                # 개선된 필터링 함수 호출
                                df_filtered, filter_message = improved_filter_data(df, filter_year)
                            
                            st.session_state.df_filtered = df_filtered.copy()
                            progress_bar.progress(33)
                            
                            # 필터링 결과 표시
                            status_text.text("2/3 단계: 필터링 결과 확인 중...")
                            st.subheader("⚙️ 필터링 세부 정보")
                            st.info(filter_message)
                            
                            if df_filtered.empty:
                                st.warning("⚠️ 필터링 후 남은 데이터가 없습니다.")
                                progress_bar.progress(100)
                                status_text.text("처리 완료: 남은 데이터가 없습니다.")
                            else:
                                st.write(f"- 원본 데이터 행 수: {len(df)}")
                                st.write(f"- 필터링 후 행 수: {len(df_filtered)}")
                                
                                # 필터링된 데이터 미리보기
                                st.subheader("🔄 필터링된 데이터")
                                st.dataframe(df_filtered, height=300)
                                
                                # 3단계: 새 포맷으로 변환
                                status_text.text("3/3 단계: 새 포맷으로 변환 중...")
                                progress_bar.progress(66)
                                
                                try:
                                    # 총평점 열 확인 및 처리
                                    if '총평점' not in df_filtered.columns:
                                        st.warning("'총평점' 열이 없습니다. 기본값 0을 사용합니다.")
                                        df_filtered['총평점'] = 0
                                    
                                    # 대상구분 열 확인 및 처리
                                    if '대상구분' not in df_filtered.columns:
                                        st.warning("'대상구분' 열이 없습니다. 기본값을 사용합니다.")
                                        df_filtered['대상구분'] = "대상"
                                    
                                    # 새 포맷으로 변환 (개선된 함수 사용)
                                    with log_container:
                                        st.subheader("📋 변환 과정 로그")
                                        new_format_df = improved_transform_to_new_format(df_filtered)
                                    
                                    st.session_state.df_transformed = new_format_df.copy()
                                    
                                    # 변환 완료
                                    progress_bar.progress(100)
                                    status_text.text("✅ 처리 완료")
                                    
                                    # 변환된 데이터가 비어있는지 확인
                                    if new_format_df.empty:
                                        st.error("❌ 변환 결과가 비어 있습니다. 로그를 확인하여 원인을 파악하세요.")
                                    else:
                                        # 변환된 데이터 미리보기
                                        st.subheader("🔄 변환된 데이터")
                                        st.dataframe(new_format_df, height=300)
                                        
                                        # 엑셀 파일 생성 및 다운로드 옵션
                                        st.subheader("📥 파일 다운로드")
                                        
                                        # 엑셀 생성 옵션
                                        excel_option = st.radio(
                                            "엑셀 파일 형식:",
                                            ["기본 형식", 
                                            # "테두리 서식 적용"
                                             ]
                                        )
                                        
                                        # 엑셀 파일 생성
                                        buffer = io.BytesIO()
                                        
                                        if excel_option == "기본 형식":
                                            # 기본 형식 - 간단한 엑셀 변환
                                            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                                                new_format_df.to_excel(writer, index=False, sheet_name='Data')
                                        else:
                                            # 테두리 서식 적용 - 상세 서식 지정
                                            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                                                new_format_df.to_excel(writer, index=False, sheet_name='Data')
                                                
                                                # 워크시트와 workbook 객체 가져오기
                                                workbook = writer.book
                                                worksheet = writer.sheets['Data']
                                                
                                                # 헤더 서식 설정
                                                header_fmt = workbook.add_format({
                                                    'bold': True, 
                                                    'border': 1,
                                                    'bg_color': '#D9E1F2',
                                                    'align': 'center',
                                                    'valign': 'vcenter'
                                                })
                                                
                                                # 데이터 셀 서식 설정
                                                cell_fmt = workbook.add_format({'border': 1})
                                                
                                                # 헤더에 서식 적용
                                                for col_num, col_name in enumerate(new_format_df.columns):
                                                    worksheet.write(0, col_num, col_name, header_fmt)
                                                
                                                # 데이터 셀에 서식 적용
                                                for row_num in range(1, len(new_format_df) + 1):
                                                    for col_num in range(len(new_format_df.columns)):
                                                        # 조건부 서식을 사용하여 모든 셀에 테두리 적용
                                                        worksheet.conditional_format(
                                                            row_num, col_num, row_num, col_num,
                                                            {'type': 'no_blanks', 'format': cell_fmt}
                                                        )
                                                        worksheet.conditional_format(
                                                            row_num, col_num, row_num, col_num,
                                                            {'type': 'blanks', 'format': cell_fmt}
                                                        )
                                                
                                                # 열 너비 조정
                                                worksheet.set_column('A:A', 5)   # 연번
                                                worksheet.set_column('B:B', 10)  # 성명
                                                worksheet.set_column('C:C', 10)  # 면허번호
                                                worksheet.set_column('D:D', 5)   # 직종
                                                worksheet.set_column('E:E', 12)  # 보수교육 이수년도
                                                worksheet.set_column('F:F', 12)  # 보수교육 이수시간
                                                worksheet.set_column('G:G', 15)  # 자기 종사지구분
                                                worksheet.set_column('H:H', 15)  # 필수교육 이수시간
                                                worksheet.set_column('I:I', 15)  # 회원 아이디
                                        
                                        buffer.seek(0)
                                        
                                        # 파일명 생성
                                        today = datetime.now().strftime("%m%d")
                                        download_filename = f"Data {today} 교육이수자.xlsx"
                                        
                                        # 다운로드 버튼
                                        st.download_button(
                                            label="📥 변환된 파일 다운로드",
                                            data=buffer,
                                            file_name=download_filename,
                                            mime="application/vnd.ms-excel"
                                        )
                                        
                                        st.success(f"✅ 변환 작업 완료! 총 {len(new_format_df)}개 행이 변환되었습니다.")
                                    
                                except Exception as e:
                                    st.error(f"❌ 데이터 변환 중 오류 발생: {e}")
                                    st.error(traceback.format_exc())
                                    progress_bar.progress(100)
                                    status_text.text("❌ 처리 실패")
                        
        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")
            st.error(traceback.format_exc())