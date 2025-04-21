import streamlit as st
import pandas as pd
import io
from datetime import datetime

def convert_excel_format(input_file):
    """
    엑셀 파일 형식을 변환하는 함수
    
    Args:
        input_file: 입력 엑셀 파일 객체
    
    Returns:
        tuple: (변환된 데이터프레임, 확인 결과 문자열)
    """
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

def create_excel_download(df, confirmation_result):
    """
    데이터프레임을 엑셀 파일로 다운로드할 수 있게 하는 함수
    
    Args:
        df: 데이터프레임
        confirmation_result: 확인 결과 문자열
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    buffer.seek(0)

    today = datetime.now().strftime("%m%d")
    download_filename = f"Data {today} {confirmation_result}.xlsx"

    st.download_button(
        label="📥 변환된 파일 다운로드",
        data=buffer,
        file_name=download_filename,
        mime="application/vnd.ms-excel"
    )

    st.success(f"총 {len(df)}개 행 변환 완료!")
    st.info(f"파일명: {download_filename}")
    if '이름' in df.columns:
        st.write(f"- 고유 사용자 수: {df['이름'].nunique()}")

def load_excel_file(file):
    """엑셀 파일을 로드하는 함수"""
    try:
        # 파일 확장자 확인
        file_extension = file.name.split('.')[-1].lower()
        
        # 확장자에 따라 엔진 지정
        if file_extension == 'xlsx':
            engine = 'openpyxl'
        elif file_extension == 'xls':
            engine = 'xlrd'
        else:
            st.error(f"지원하지 않는 파일 형식입니다: {file_extension}")
            return None
        
        # 엔진을 명시적으로 지정하여 파일 로드
        return pd.read_excel(file, engine=engine)
    except Exception as e:
        st.error(f"파일 로드 중 오류가 발생했습니다: {e}")
        return None

def dataframe_to_excel(df):
    """판다스 데이터프레임을 엑셀 파일로 변환하는 함수"""
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.close()
    processed_data = output.getvalue()
    return processed_data