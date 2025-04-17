import streamlit as st
import pandas as pd
from processors.excel_processor import convert_excel_format, create_excel_download

def show_exemption_tab():
    """면제/유예/비대상 탭을 표시하는 함수"""
    
    st.markdown("### 2. 면제/유예/비대상")
    
    # 파일 업로드 UI
    uploaded_excel = st.file_uploader("📥 면제/유예/비대상 .xlsx 또는 .xls 파일 업로드", type=["xlsx", "xls"], key="excel")

    if uploaded_excel:
        # 미리보기 표시
        df_preview = pd.read_excel(uploaded_excel)
        st.dataframe(df_preview.head())

        # 변환 버튼
        if st.button("변환하기"):
            # 파일 포인터 초기화
            uploaded_excel.seek(0)
            
            # 엑셀 파일 형식 변환
            converted_df, confirmation_result = convert_excel_format(uploaded_excel)

            if converted_df is not None:
                # 변환된 데이터 표시
                st.subheader("🔍 변환된 데이터")
                st.dataframe(converted_df.head(10))
                
                # 엑셀 다운로드 버튼 생성
                create_excel_download(converted_df, confirmation_result)