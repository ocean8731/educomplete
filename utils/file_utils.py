import pandas as pd
import io
from bs4 import BeautifulSoup

def decode_bytes(file_bytes):
    """
    파일 바이트를 적절한 인코딩으로 디코딩하는 함수
    
    Args:
        file_bytes: 디코딩할 파일 바이트
    
    Returns:
        str: 디코딩된 문자열
    
    Raises:
        UnicodeDecodeError: 모든 인코딩 시도 실패 시
    """
    for encoding in ['utf-8', 'euc-kr', 'cp949']:
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("❌ 파일 인코딩을 확인할 수 없습니다.")

def parse_html_xls(file_bytes, use_first_row_as_header):
    """
    HTML .xls 파일을 파싱하는 함수
    
    Args:
        file_bytes: 파일 바이트
        use_first_row_as_header: 첫 행을 헤더로 사용할지 여부
    
    Returns:
        DataFrame: 파싱된 데이터프레임
    
    Raises:
        ValueError: 테이블을 찾을 수 없는 경우
    """
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

def df_to_xlsx_bytes(df):
    """
    데이터프레임을 엑셀 바이트로 변환하는 함수
    
    Args:
        df: 변환할 데이터프레임
    
    Returns:
        BytesIO: 엑셀 파일 바이트
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    output.seek(0)
    return output