import os
import streamlit as st

from donation_report_pipeline import run_report_pipeline


def render_report_tab():
    st.subheader("회보서")

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    DATA_DIR = os.path.join(BASE_DIR, "data")
    INPUT_DIR = os.path.join(DATA_DIR, "input")
    OUTPUT_DIR = os.path.join(DATA_DIR, "output")

    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    REPORT_TEMPLATE_PATH = os.path.join(INPUT_DIR, "후원회_회보서_정리.xlsx")

    if not os.path.exists(REPORT_TEMPLATE_PATH):
        st.error("기본 파일이 없습니다. data/input 폴더의 후원회_회보서_정리.xlsx 파일을 확인해주세요.")
        return

    uploaded_file = st.file_uploader(
        "은행내역 파일 업로드",
        type=["xlsx", "xls", "csv"],
        key="report_upload_file"
    )

    if uploaded_file:
        st.success("은행내역 파일 업로드 완료")

        if st.button("회보서 파일 생성", key="report_run_btn"):
            try:
                output_path, df_result = run_report_pipeline(
                    uploaded_file=uploaded_file,
                    template_path=REPORT_TEMPLATE_PATH,
                    output_dir=OUTPUT_DIR
                )

                st.success("회보서 파일 생성 완료")

                with open(output_path, "rb") as f:
                    st.download_button(
                        label="결과 파일 다운로드",
                        data=f,
                        file_name=os.path.basename(output_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="report_download_btn"
                    )

            except Exception as e:
                st.error(f"오류 발생: {e}")
