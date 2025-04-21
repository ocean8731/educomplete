import streamlit as st
import pandas as pd
import numpy as np
import traceback
import io
from datetime import datetime

def improved_transform_to_new_format(df_filtered):
    """
    필터링된 데이터를 새 포맷으로 변환하는 개선된 함수
    
    Args:
        df_filtered: 필터링된 데이터프레임
        
    Returns:
        DataFrame: 변환된 데이터프레임
    """
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
            
        # 새 행 구성 - 딕셔너리로 생성 (정수형으로 변환)
        new_row = {
            '연번': len(data) + 1,
            '성명': row['성명'],
            '면허번호': int(pd.to_numeric(row['면허번호'], errors='coerce') or 0),
            '직종': 15,
            '보수교육 이수년도': int(pd.to_numeric(row['이수년도'], errors='coerce') or 0),
            '보수교육 이수시간': int(pd.to_numeric(row['총평점'], errors='coerce') or 0) if '총평점' in row and not pd.isna(row['총평점']) else 0,
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

def create_completion_excel(new_format_df, excel_option="기본 형식"):
    """
    보수교육 완료자 명단 엑셀 파일 생성
    
    Args:
        new_format_df: 변환된 데이터프레임
        excel_option: 엑셀 파일 형식 옵션
    
    Returns:
        tuple: (BytesIO 객체, 파일명)
    """
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
    
    return buffer, download_filename

def compare_license_data(df1, df2):
    """첫 번째 파일(기준)과 두 번째 파일(비교대상)을 비교"""
    # 두 파일 모두에 있는 데이터 비교
    common_data = pd.merge(df1, df2, on='면허번호', how='inner', suffixes=('_1', '_2'))
    
    # 상태 열 추가
    common_data['상태'] = np.where(
        common_data['성명_1'] == common_data['성명_2'], 
        '일치', 
        '불일치'
    )
    
    # 불일치하는 경우 차이점 표시
    common_data['차이점'] = ''
    mask = common_data['상태'] == '불일치'
    common_data.loc[mask, '차이점'] = '성명 불일치'
    
    # 첫 번째 파일에만 있는 데이터 찾기 (두 번째 파일에는 없는 데이터)
    only_in_first = pd.merge(
        df1, 
        df2[['면허번호']], 
        on='면허번호', 
        how='left', 
        indicator=True
    )
    only_in_first = only_in_first[only_in_first['_merge'] == 'left_only']
    only_in_first = only_in_first[['면허번호', '성명']]
    only_in_first['상태'] = '두 번째 파일에 누락'
    only_in_first['차이점'] = '두 번째 파일에 존재하지 않음'
    
    # 결과 집계
    total_common = len(common_data)
    matched = len(common_data[common_data['상태'] == '일치'])
    unmatched = len(common_data[common_data['상태'] == '불일치'])
    only_in_first_count = len(only_in_first)
    
    # 일치율 계산 (공통 데이터 내에서)
    match_rate = (matched / total_common) * 100 if total_common > 0 else 0
    
    return common_data, only_in_first, {
        '총 비교 레코드': total_common,
        '일치 레코드': matched,
        '불일치 레코드': unmatched,
        '첫 번째 파일에만 있는 레코드': only_in_first_count,
        '일치율': match_rate
    }