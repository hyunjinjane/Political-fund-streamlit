import streamlit as st

def render_center_tab():
    st.subheader("센터 입력용")

    uploaded_file = st.file_uploader(
        "센터 입력용 파일 업로드",
        type=["xlsx", "xls", "csv"],
        key="center_file"
    )

    if uploaded_file:
        st.success("센터 입력용 파일 업로드 완료")
        # 센터 입력용 처리 코드