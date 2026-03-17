import streamlit as st

def render_report_tab():
    st.subheader("회보서")

    uploaded_file = st.file_uploader(
        "회보서 파일 업로드",
        type=["xlsx", "xls", "csv"],
        key="report_file"
    )

    if uploaded_file:
        st.success("회보서 파일 업로드 완료")
        # 회보서 처리 코드