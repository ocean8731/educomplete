import streamlit as st
import pandas as pd
import io
import zipfile
from utils.file_utils import parse_html_xls, df_to_xlsx_bytes
from utils.data_utils import transform_and_split
from utils.date_utils import get_last_workday_str, get_output_filename

def process_html_file(uploaded_file, use_first_row_as_header):
    """
    HTML .xls 파일을 처리하는 함수
    
    Args:
        uploaded_file: 업로드된 파일 객체
        use_first_row_as_header: 첫 행을 헤더로 사용할지 여부
    
    Returns:
        bool: 처리 성공 여부
    """
    try:
        # 파일 읽기 및 파싱
        df = parse_html_xls(uploaded_file.read(), use_first_row_as_header)
        
        # '소속' 컬럼 제거 (있는 경우)
        if '소속' in df.columns:
            df = df.drop(columns=['소속'])

        # 데이터 미리보기 표시
        st.dataframe(df.astype(str))
        
        # 데이터 변환 및 분할
        output_dfs = transform_and_split(df)
        
        # 출력 파일 이름 생성을 위한 날짜 문자열
        today_str = get_last_workday_str()
        
        # ZIP 파일 생성
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for i, df_part in enumerate(output_dfs):
                xlsx_bytes = df_to_xlsx_bytes(df_part)
                filename = get_output_filename(today_str, None if i == 0 else i + 1)
                zip_file.writestr(filename, xlsx_bytes.read())

        # 다운로드 버튼 생성
        zip_buffer.seek(0)
        st.download_button(
            label="📦 ZIP 파일 다운로드",
            data=zip_buffer,
            file_name=f"({today_str}) 현장 보수교육 대체 사이버 교육 이수자 리스트.zip",
            mime="application/zip"
        )
        
        return True
    except Exception as e:
        st.error(f"❌ 오류 발생: {e}")
        return False