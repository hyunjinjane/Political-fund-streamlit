import streamlit as st

from donation_program_page import render_program_tab
from donation_center_page import render_center_tab
from donation_report_page import render_report_tab

st.set_page_config(page_title="후원금 정리", page_icon="💰", layout="wide")

st.title("후원금 정리")

tab1, tab2, tab3, tab4 = st.tabs(["프로그램 입력용", "센터 입력용", "회보서", "연말정산 업로드용"])

with tab1:
    render_program_tab()

with tab2:
    render_center_tab()

with tab3:
    render_report_tab()

