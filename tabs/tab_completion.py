import streamlit as st
import pandas as pd
from datetime import datetime
import traceback
from utils.file_utils import parse_html_xls
from utils.data_utils import improved_filter_data
from processors.transform_processor import improved_transform_to_new_format, create_completion_excel_chunks

def show_completion_tab():
    """보수교육 완료자 명단등록 탭을 표시하는 함수"""

    st.markdown("### 3. 보수교육 완료자 명단등록")

    # 상태 관리용 변수
    if 'df_original' not in st.session_state:
        st.session_state.df_original = None
    if 'df_filtered' not in st.session_state:
        st.session_state.df_filtered = None
    if 'df_transformed' not in st.session_state:
        st.session_state.df_transformed = None
    if 'last_uploaded_file' not in st.session_state:
        st.session_state.last_uploaded_file = None

    # 파일 업로드
    uploaded_html_filter = st.file_uploader("📤 필터링할 .xls HTML 파일 업로드", type=["xls"], key="html_filter")

    # 새 파일 업로드 시 이전 변환 결과 초기화
    if uploaded_html_filter is not None:
        if st.session_state.last_uploaded_file != uploaded_html_filter.name:
            st.session_state.df_transformed = None
            st.session_state.df_filtered = None
            st.session_state.last_uploaded_file = uploaded_html_filter.name

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
        st.subheader("🔧 고급 데이터 검증 옵션")
        col1, col2 = st.columns(2)
        with col1:
            remove_empty_rows = st.checkbox("빈 행 자동 제거", value=True,
                                           help="모든 값이 비어있거나 NA인 행을 자동으로 제거합니다.")
        with col2:
            validate_required_fields = st.checkbox("필수 필드 검증", value=True,
                                                help="면허번호, 성명, 이수년도 필드가 비어있는 행을 제외합니다.")
    else:
        remove_empty_rows = True
        validate_required_fields = True

    if uploaded_html_filter:
        try:
            with st.spinner("파일 파싱 중..."):
                file_content = uploaded_html_filter.read()
                if len(file_content) == 0:
                    st.error("❌ 빈 파일입니다.")
                else:
                    df = parse_html_xls(file_content, use_first_row)
                    st.session_state.df_original = df.copy()

                    st.success("✅ 파일 파싱 완료")

                    st.subheader("🔍 원본 데이터 컬럼 정보")
                    col_info = pd.DataFrame({
                        "컬럼명": df.columns,
                        "데이터 타입": df.dtypes.astype(str),
                        "샘플값": [str(df[col].iloc[0]) if len(df) > 0 else "" for col in df.columns],
                        "고유값 수": [df[col].nunique() for col in df.columns],
                        "결측값 수": [df[col].isna().sum() for col in df.columns]
                    })
                    st.dataframe(col_info)

                    st.subheader("🔍 업로드된 데이터 미리보기")
                    st.dataframe(df, height=300)

                    required_columns = ['면허번호', '성명', '이수년도']
                    missing = [col for col in required_columns if col not in df.columns]

                    if missing:
                        st.error(f"❌ 필요한 컬럼이 누락되었습니다: {missing}")
                        st.info("업로드된 파일의 컬럼명을 확인하세요. 첫 번째 행을 컬럼명으로 사용하는 옵션을 조정해보세요.")
                    else:
                        # --- 필터링 및 변환 버튼 ---
                        if st.button("필터링 및 변환 실행"):
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            log_container = st.expander("상세 로그 보기", expanded=False)

                            status_text.text("1/3 단계: 데이터 필터링 중...")
                            with log_container:
                                st.subheader("📋 필터링 과정 로그")
                                df_filtered, filter_message = improved_filter_data(df, filter_year)

                            st.session_state.df_filtered = df_filtered.copy()
                            progress_bar.progress(33)

                            status_text.text("2/3 단계: 필터링 결과 확인 중...")

                            if isinstance(df_filtered, pd.DataFrame) and df_filtered.empty:
                                st.warning("⚠️ 필터링 후 남은 데이터가 없습니다.")
                                progress_bar.progress(100)
                                status_text.text("처리 완료: 남은 데이터가 없습니다.")
                            else:
                                status_text.text("3/3 단계: 새 포맷으로 변환 중...")
                                progress_bar.progress(66)

                                try:
                                    if '총평점' not in df_filtered.columns:
                                        df_filtered['총평점'] = 0
                                    if '대상구분' not in df_filtered.columns:
                                        df_filtered['대상구분'] = "대상"

                                    with log_container:
                                        st.subheader("📋 변환 과정 로그")
                                        new_format_df = improved_transform_to_new_format(df_filtered)

                                    st.session_state.df_transformed = new_format_df.copy()
                                    progress_bar.progress(100)
                                    status_text.text("✅ 처리 완료")

                                except Exception as e:
                                    st.error(f"❌ 데이터 변환 중 오류 발생: {e}")
                                    st.error(traceback.format_exc())
                                    progress_bar.progress(100)
                                    status_text.text("❌ 처리 실패")

        except Exception as e:
            st.error(f"❌ 오류 발생: {e}")
            st.error(traceback.format_exc())

    # --- 다운로드 섹션: session_state 기반으로 항상 유지 ---
    if st.session_state.df_transformed is not None:
        new_format_df = st.session_state.df_transformed

        if new_format_df.empty:
            st.error("❌ 변환 결과가 비어 있습니다. 로그를 확인하여 원인을 파악하세요.")
        else:
            st.subheader("🔄 변환된 데이터")
            st.dataframe(new_format_df, height=300)

            st.subheader("📥 파일 다운로드")

            split_option = st.selectbox(
                "파일 분할 방식",
                options=["분할 없음", "100개 단위로 분할", "1000개 단위로 분할"],
                index=1
            )

            chunk_size_map = {
                "분할 없음": None,
                "100개 단위로 분할": 100,
                "1000개 단위로 분할": 1000,
            }
            chunk_size = chunk_size_map[split_option]

            if chunk_size is None:
                chunks = create_completion_excel_chunks(new_format_df, "기본 형식", chunk_size=len(new_format_df))
            else:
                chunks = create_completion_excel_chunks(new_format_df, "기본 형식", chunk_size=chunk_size)

            if len(chunks) == 1:
                buffer, download_filename = chunks[0]
                st.download_button(
                    label="📥 변환된 파일 다운로드",
                    data=buffer,
                    file_name=download_filename,
                    mime="application/vnd.ms-excel"
                )
            else:
                st.info(f"데이터가 {chunk_size}개를 초과하여 총 {len(chunks)}개 파일로 분리됩니다.")
                for idx, (buffer, download_filename) in enumerate(chunks, start=1):
                    st.download_button(
                        label=f"📥 파일 {idx} 다운로드 ({download_filename})",
                        data=buffer,
                        file_name=download_filename,
                        mime="application/vnd.ms-excel",
                        key=f"download_chunk_{idx}"
                    )

            st.success(f"✅ 변환 작업 완료! 총 {len(new_format_df)}개 행이 변환되었습니다.")
