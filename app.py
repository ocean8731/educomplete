import streamlit as st
import pandas as pd
import io
import zipfile
import xlwt
import holidays
from datetime import datetime, date, timedelta
from bs4 import BeautifulSoup

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

# 탭 구성
tab1, tab2 = st.tabs(["💡 현장대체온라인교육", "📊 면제/유예/비대상"])

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
