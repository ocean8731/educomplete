import streamlit as st
import pandas as pd
from bs4 import BeautifulSoup
import io
import xlwt
import zipfile
import holidays
from datetime import date, timedelta

st.set_page_config(page_title="현장대체온라인교육", layout="centered")
st.title("📄 현장대체온라인교육")
# 사용 안내
st.markdown("""
### 변환 내용
- 페이지에서 다운로드 받는 액셀 파일에 이 데이터가 대상인지/비대상인지가 없습니다.
- 대상/비대상만 파악하셔서 파일이름을 수정해 주세요.            
""")
st.caption("변환할 .xls 파일을 업로드하세요")

uploaded_file = st.file_uploader("📤 파일 업로드", type=["xls"])

def get_last_workday_str():
    kr_holidays = holidays.KR(years=date.today().year)
    today = date.today()
    day = today - timedelta(days=1)  # 기본 하루 전

    # 한국 공휴일 또는 주말이면 더 거슬러 올라감
    while day in kr_holidays or day.weekday() >= 5:
        day -= timedelta(days=1)

    return day.strftime("%y.%m.%d")

# 오늘 날짜 형식 지정
today_str = get_last_workday_str()

def get_output_filename(index=None):
    base = f"({today_str}_대상) 현장 보수교육 대체 사이버 교육 이수자 리스트"
    return f"{base}.xls" if index is None else f"{base}_{index}.xls"

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

def transform_and_split(df):
    required = ['면허번호', '성명', '이수년도']
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"❌ 필요한 컬럼이 누락되었습니다: {missing}")

    df = df[required].copy()
    df['이수년도'] = df['이수년도'].astype(str)
    grouped = df.groupby(['면허번호', '성명'])

    file_data = []
    from collections import defaultdict
    index_map = defaultdict(int)

    for (lic, name), group in grouped:
        group = group.sort_values(by='이수년도')
        for i, (_, row) in enumerate(group.iterrows()):
            target_index = index_map[(lic, name)]  # 현재 이 사람의 몇 번째 데이터인지
            if len(file_data) <= target_index:
                file_data.append([])
            file_data[target_index].append(row)
            index_map[(lic, name)] += 1

    # 순번 붙이고 소속 제거
    output_dfs = []
    for file_rows in file_data:
        out_df = pd.DataFrame(file_rows).reset_index(drop=True)
        out_df.insert(0, "순번", range(1, len(out_df) + 1))
        output_dfs.append(out_df)

    return output_dfs

def df_to_xls_bytes(df):
    output = io.BytesIO()
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Data")
    for col_idx, col_name in enumerate(df.columns):
        ws.write(0, col_idx, col_name)
    for row_idx, row in enumerate(df.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row):
            ws.write(row_idx, col_idx, value)
    wb.save(output)
    output.seek(0)
    return output



if uploaded_file:
    use_first_row_as_header = st.checkbox("🔠 첫 번째 행을 컬럼명으로 사용할까요?", value=True)

    try:
        file_bytes = uploaded_file.read()
        df = parse_html_xls(file_bytes, use_first_row_as_header)

        # 소속 제거
        if '소속' in df.columns:
            df = df.drop(columns=['소속'])

        st.subheader("🔍 원본 테이블 미리보기")
        st.dataframe(df.astype(str))

        output_dfs = transform_and_split(df)

        st.subheader("✅ 변환된 파일 수: " + str(len(output_dfs)))

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for i, df_part in enumerate(output_dfs):
                xls_bytes = df_to_xls_bytes(df_part)
                filename = get_output_filename(None if i == 0 else i + 1)
                zip_file.writestr(filename, xls_bytes.read())

        zip_buffer.seek(0)
        st.download_button(
            label="📦 모든 변환 파일 ZIP으로 다운로드",
            data=zip_buffer,
            file_name=f"({today_str}_) 사이버 교육 이수자 리스트.zip",
            mime="application/zip"
        )

    except Exception as e:
        st.error(f"❌ 오류 발생: {e}")
