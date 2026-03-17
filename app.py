import streamlit as st

st.set_page_config(
    page_title="업무 자동화",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("업무 자동화")
st.markdown(
    """
왼쪽 사이드바에서 원하는 기능을 선택하세요.

### 사용 가능한 기능
- 정치자금 정리
- 후원금 정리
"""
)

st.info("각 기능은 왼쪽 페이지 메뉴에서 선택할 수 있습니다.")
