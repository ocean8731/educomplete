import streamlit as st
from processors.html_processor import process_html_file

def show_online_education_tab():
    """현장대체온라인교육 탭을 표시하는 함수"""
    
    st.markdown("### 현장대체온라인교육")
    
    # 파일 업로드 UI
    uploaded_html = st.file_uploader("📤 현장대체온라인교육 .xls 파일 업로드", type=["xls"], key="html")

    if uploaded_html:
        # 옵션 설정
        use_first_row_as_header = st.checkbox("🔠 첫 번째 행을 컬럼명으로 사용할까요?", value=True)
        
        # 파일 처리
        process_html_file(uploaded_html, use_first_row_as_header)