import streamlit as st
from tabs.tab_online_education import show_online_education_tab
from tabs.tab_exemption import show_exemption_tab
from tabs.tab_completion import show_completion_tab
from tabs.tab_license_verification import show_license_verification_tab  # 새로 추가한 탭 임포트

# 페이지 설정
st.set_page_config(page_title="엑셀 자동화 도구", layout="wide")
st.title("📄 엑셀 자동화 도구")

# 탭 구성 (새 탭 추가)
tab1, tab2, tab3, tab4 = st.tabs(["💡 현장대체온라인교육", "📊 면제/유예/비대상", "🔄 보수교육 완료자 명단등록", "🔍 면허번호/성명 검증"])

# 각 탭 컨텐츠 표시
with tab1:
    show_online_education_tab()

with tab2:
    show_exemption_tab()

with tab3:
    show_completion_tab()

# 새로운 탭 추가
with tab4:
    show_license_verification_tab()