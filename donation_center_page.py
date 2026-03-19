import os
import json
import pandas as pd
import streamlit as st

from donation_center_pipeline import run_center_pipeline


DEFAULT_CENTER_CONFIG = {
    "defaults": {
        "주민등록번호(*)": "991231",
        "연락처(*)": "010-1234-1234",
        "직업분류(*)": "기타",
        "직업": "",
        "이메일": ""
    },
    "address": {
        "우편번호(*)": "",
        "주소(*)": "",
        "상세주소(*)": ""
    }
}


def _config_to_df(config_dict):
    defaults = config_dict.get("defaults", {})
    address = config_dict.get("address", {})

    return pd.DataFrame([{
        "주민등록번호(*)": defaults.get("주민등록번호(*)", ""),
        "연락처(*)": defaults.get("연락처(*)", ""),
        "직업분류(*)": defaults.get("직업분류(*)", ""),
        "직업": defaults.get("직업", ""),
        "이메일": defaults.get("이메일", ""),
        "우편번호(*)": address.get("우편번호(*)", ""),
        "주소(*)": address.get("주소(*)", ""),
        "상세주소(*)": address.get("상세주소(*)", "")
    }])


def _df_to_config(df):
    if df is None or df.empty:
        return DEFAULT_CENTER_CONFIG.copy()

    row = df.iloc[0]

    def val(col):
        value = row.get(col, "")
        return "" if pd.isna(value) else str(value)

    return {
        "defaults": {
            "주민등록번호(*)": val("주민등록번호(*)"),
            "연락처(*)": val("연락처(*)"),
            "직업분류(*)": val("직업분류(*)"),
            "직업": val("직업"),
            "이메일": val("이메일")
        },
        "address": {
            "우편번호(*)": val("우편번호(*)"),
            "주소(*)": val("주소(*)"),
            "상세주소(*)": val("상세주소(*)")
        }
    }


def render_center_tab():
    st.subheader("센터 입력용")

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    DATA_DIR = os.path.join(BASE_DIR, "data")
    INPUT_DIR = os.path.join(DATA_DIR, "input")
    OUTPUT_DIR = os.path.join(DATA_DIR, "output")

    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    CENTER_TEMPLATE_PATH = os.path.join(INPUT_DIR, "후원회_센터_등록.xlsx")

    if not os.path.exists(CENTER_TEMPLATE_PATH):
        st.error("기본 파일이 없습니다. data/input 폴더의 후원회_센터_등록.xlsx 파일을 확인해주세요.")
        return

    if "center_config_data" not in st.session_state:
        st.session_state.center_config_data = json.loads(json.dumps(DEFAULT_CENTER_CONFIG))

    uploaded_file = st.file_uploader(
        "은행내역 파일 업로드",
        type=["xlsx", "xls", "csv"],
        key="center_upload_file"
    )

    input_mode = st.radio(
        "설정 입력 방식",
        ["직접 입력", "JSON 업로드"],
        horizontal=True,
        key="center_config_mode"
    )

    st.markdown("### 센터 설정")

    if input_mode == "직접 입력":
        current_df = _config_to_df(st.session_state.center_config_data)

        edited_df = st.data_editor(
            current_df,
            num_rows="fixed",
            use_container_width=True,
            key="center_config_editor_direct"
        )

        col1, col2 = st.columns(2)

        with col1:
            if st.button("설정 적용하기", key="center_apply_direct_btn"):
                st.session_state.center_config_data = _df_to_config(edited_df)
                st.success("설정이 적용되었습니다.")

        with col2:
            json_bytes = json.dumps(
                _df_to_config(edited_df),
                ensure_ascii=False,
                indent=2
            ).encode("utf-8")

            st.download_button(
                label="수정된 JSON 다운로드",
                data=json_bytes,
                file_name="센터설정.json",
                mime="application/json",
                key="center_download_direct_json"
            )

    else:
        uploaded_json = st.file_uploader(
            "설정 JSON 업로드",
            type=["json"],
            key="center_json_upload"
        )

        if uploaded_json is not None:
            try:
                uploaded_config = json.load(uploaded_json)
                uploaded_df = _config_to_df(uploaded_config)

                st.write("업로드한 설정")
                edited_uploaded_df = st.data_editor(
                    uploaded_df,
                    num_rows="fixed",
                    use_container_width=True,
                    key="center_config_editor_json"
                )

                col1, col2 = st.columns(2)

                with col1:
                    if st.button("업로드한 설정 적용하기", key="center_apply_json_btn"):
                        st.session_state.center_config_data = _df_to_config(edited_uploaded_df)
                        st.success("업로드한 설정이 적용되었습니다.")

                with col2:
                    edited_json_bytes = json.dumps(
                        _df_to_config(edited_uploaded_df),
                        ensure_ascii=False,
                        indent=2
                    ).encode("utf-8")

                    st.download_button(
                        label="수정된 JSON 다운로드",
                        data=edited_json_bytes,
                        file_name="센터설정_수정본.json",
                        mime="application/json",
                        key="center_download_uploaded_json"
                    )

            except Exception as e:
                st.error(f"JSON 읽기 오류: {e}")

        with st.expander("JSON 예시"):
            st.dataframe(
                _config_to_df(DEFAULT_CENTER_CONFIG),
                use_container_width=True
            )

    st.markdown("### 현재 적용된 설정")
    st.dataframe(
        _config_to_df(st.session_state.center_config_data),
        use_container_width=True
    )

    if uploaded_file:
        st.success("은행내역 파일 업로드 완료")

        if st.button("센터 입력용 파일 생성", key="center_run_btn"):
            try:
               output_path, df_result = run_center_pipeline(
                 uploaded_file=uploaded_file,
                 template_path=CENTER_TEMPLATE_PATH,
                 output_dir=OUTPUT_DIR,
                 config_source=st.session_state.center_config_data
                )

               st.success("파일 생성 완료")

               with open(output_path, "rb") as f:
                    st.download_button(
                       label="결과 파일 다운로드",
                       data=f,
                       file_name=os.path.basename(output_path),
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key="center_download_btn"
                    )

            except Exception as e:
                st.error(f"오류 발생: {e}")




