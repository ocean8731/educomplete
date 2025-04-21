import streamlit as st
from utils.data_utils import preprocess_data, get_data_overview
from processors.excel_processor import load_excel_file, dataframe_to_excel
from processors.transform_processor import compare_license_data

def show_license_verification_tab():
    st.header("📋 면허번호/성명 데이터 일치 검증")
    st.write("두 엑셀 파일의 면허번호와 성명 데이터를 비교합니다.")
    
    # 파일 업로드
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("이수자 데이터")
        file1 = st.file_uploader("이수자 데이터를 넣어주세요", type=['xlsx', 'xls'], key="license_file1")
    
    with col2:
        st.subheader("회원 데이터")
        file2 = st.file_uploader("회원 데이터를 넣어주세요. 회원 데이터를 넣기 전에, 다운로드 받은 원본 파일을 다른 이름으로 저장 -> 파일 형식을 Excel 통합 문서 로 바꿔주세요. ", type=['xlsx', 'xls'], key="license_file2")
    
    if file1 is not None and file2 is not None:
        # 엑셀 파일 로드
        df1 = load_excel_file(file1)
        df2 = load_excel_file(file2)
        
        if df1 is not None and df2 is not None:
            st.success("두 파일이 성공적으로 로드되었습니다!")
            
            # 열 선택 섹션
            st.subheader("열 선택")
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("**이수자 데이터**")
                license_col1 = st.selectbox("면허번호 열 선택 (파일 1)", options=df1.columns, key="license_col1")
                name_col1 = st.selectbox("성명 열 선택 (파일 1)", options=df1.columns, key="name_col1")
            
            with col2:
                st.markdown("**회원 데이터**")
                license_col2 = st.selectbox("면허번호 열 선택 (파일 2)", options=df2.columns, key="license_col2")
                name_col2 = st.selectbox("성명 열 선택 (파일 2)", options=df2.columns, key="name_col2")
            
            # 데이터 전처리
            df1_processed = preprocess_data(df1, license_col1, name_col1)
            df2_processed = preprocess_data(df2, license_col2, name_col2)
            
            # 데이터 개요 계산
            overview, common_data_preview, only_in_first = get_data_overview(df1_processed, df2_processed)
            
            # 불일치 데이터 추출 (미리보기용)
            unmatched_preview = common_data_preview[common_data_preview['상태'] == '불일치']
            
            # 데이터 개요 표시
            st.subheader("데이터 개요")
            
            # 파일 레코드 수
            col1, col2 = st.columns(2)
            with col1:
                st.metric("첫 번째 파일 레코드 수", overview['first_file_count'])
            with col2:
                st.metric("두 번째 파일 레코드 수", overview['second_file_count'])
            
            # 공통 면허번호/첫 번째 파일에만 있는 레코드 수
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("공통 면허번호 수", overview['common_count'])
            with col2:
                st.metric("첫 번째 파일에만 있는 레코드 수", overview['only_in_first_count'])
            with col3:
                st.metric("첫 번째 파일 - 공통", overview['first_minus_common'])
            
            # 불일치 데이터 통계
            st.subheader("불일치 데이터 통계")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("공통 데이터 중 일치 수", overview['matched_count'])
            with col2:
                st.metric("공통 데이터 중 불일치 수", overview['unmatched_count'])
            with col3:
                st.metric("일치율", f"{overview['match_rate']:.2f}%")
            
            # 불일치 데이터 미리보기
            if len(unmatched_preview) > 0:
                st.subheader("불일치 데이터 미리보기")
                unmatched_preview_display = unmatched_preview[['면허번호', '성명_1', '성명_2']].copy()
                unmatched_preview_display.columns = ['면허번호', '첫 번째 파일 성명', '두 번째 파일 성명']
                st.dataframe(unmatched_preview_display.head(10), use_container_width=True)
                
                if len(unmatched_preview) > 10:
                    st.info(f"총 {len(unmatched_preview)}개의 불일치 데이터 중 10개만 표시됩니다. 전체 보기는 '데이터 비교 실행' 버튼을 클릭하세요.")
            
            # 비교 실행 버튼
            if st.button("데이터 비교 실행", key="compare_license_data"):
                # 데이터 비교
                common_data, only_in_first_data, stats = compare_license_data(df1_processed, df2_processed)
                
                # 결과 표시
                st.subheader("비교 결과")
                
                # 통계 표시
                col1, col2 = st.columns(2)
                col1.metric("총 비교 레코드", stats['총 비교 레코드'])
                col2.metric("일치율", f"{stats['일치율']:.2f}%")
                
                col1, col2, col3 = st.columns(3)
                col1.metric("일치 레코드", stats['일치 레코드'])
                col2.metric("불일치 레코드", stats['불일치 레코드'])
                col3.metric("첫 번째 파일에만 있는 레코드", stats['첫 번째 파일에만 있는 레코드'])
                
                # 탭으로 결과 구분
                result_tabs = st.tabs(["공통 데이터 비교", "첫 번째 파일에만 있는 데이터", "불일치 데이터만 보기"])
                
                # 공통 데이터 비교 탭
                with result_tabs[0]:
                    st.subheader("공통 데이터 비교 결과")
                    
                    # 필터 추가
                    status_filter = st.multiselect(
                        "상태 필터링",
                        options=['일치', '불일치'],
                        default=['일치', '불일치'],
                        key="status_filter"
                    )
                    
                    # 필터링된 결과 표시
                    if status_filter:
                        filtered_df = common_data[common_data['상태'].isin(status_filter)]
                        st.dataframe(filtered_df, use_container_width=True)
                        
                        # 결과 다운로드 기능 (XLSX)
                        excel_data = dataframe_to_excel(filtered_df)
                        # on_click 함수 없이 download_button 사용
                        st.download_button(
                            label="XLSX로 공통 데이터 결과 다운로드",
                            data=excel_data,
                            file_name="common_data_comparison.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_common",
                            # key_변경과 동시에 use_container_width 추가
                            use_container_width=False
                        )
                    else:
                        st.info("상태를 선택하여 결과를 필터링하세요.")
                
                # 첫 번째 파일에만 있는 데이터 탭
                with result_tabs[1]:
                    st.subheader("첫 번째 파일에만 있는 데이터 (두 번째 파일에 누락)")
                    if len(only_in_first_data) > 0:
                        st.dataframe(only_in_first_data, use_container_width=True)
                        
                        # 결과 다운로드 기능 (XLSX)
                        excel_data = dataframe_to_excel(only_in_first_data)
                        st.download_button(
                            label="XLSX로 누락 데이터 다운로드",
                            data=excel_data,
                            file_name="only_in_first_data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_missing"
                        )
                    else:
                        st.info("첫 번째 파일에만 있는 데이터가 없습니다.")
                
                # 불일치 데이터만 보기 탭
                with result_tabs[2]:
                    st.subheader("불일치 데이터만 보기")
                    unmatched_data = common_data[common_data['상태'] == '불일치']
                    
                    if len(unmatched_data) > 0:
                        st.dataframe(unmatched_data, use_container_width=True)
                        
                        # 불일치 데이터만 다운로드 기능 (XLSX)
                        excel_data = dataframe_to_excel(unmatched_data)
                        st.download_button(
                            label="XLSX로 불일치 데이터 다운로드",
                            data=excel_data,
                            file_name="unmatched_data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_unmatched"
                        )
                    else:
                        st.info("불일치하는 데이터가 없습니다.")