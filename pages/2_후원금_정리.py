import os
import json
import pandas as pd
import streamlit as st

from donation_pipeline import run_donation_pipeline
from expense_pipeline import run_expense_pipeline


st.set_page_config(page_title="후원금 정리", page_icon="💰", layout="wide")

st.title("후원금 정리")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, "data")
INPUT_DIR = os.path.join(DATA_DIR, "input")
OUTPUT_DIR = os.path.join(DATA_DIR, "output")

os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

INCOME_TEMPLATE_PATH = os.path.join(INPUT_DIR, "후원회_프로그램_수입_일괄등록.xlsx")
EXPENSE_TEMPLATE_PATH = os.path.join(INPUT_DIR, "후원회_프로그램_지출_일괄등록.xlsx")


DEFAULT_EXPENSE_CONFIG = {
    "targets": [
        {
            "지출대상자": "효성CMS",
            "생년월일(사업자번호)": "",
            "우편번호": "",
            "주 소": "",
            "상세주소": "",
            "직업(업종)": "",
            "전화번호": ""
        },
        {
            "지출대상자": "카카오페이",
            "생년월일(사업자번호)": "",
            "우편번호": "",
            "주 소": "",
            "상세주소": "",
            "직업(업종)": "",
            "전화번호": ""
        },
        {
            "지출대상자": "토스페이먼트",
            "생년월일(사업자번호)": "",
            "우편번호": "",
            "주 소": "",
            "상세주소": "",
            "직업(업종)": "",
            "전화번호": ""
        },
        {
            "지출대상자": "엔에이치엔페이코",
            "생년월일(사업자번호)": "",
            "우편번호": "",
            "주 소": "",
            "상세주소": "",
            "직업(업종)": "",
            "전화번호": ""
        },
        {
            "지출대상자": "NH농협은행",
            "생년월일(사업자번호)": "",
            "우편번호": "",
            "주 소": "",
            "상세주소": "",
            "직업(업종)": "",
            "전화번호": ""
        }
    ],
    "rules": [
        {
            "은행입력키워드": "CMS사용료",
            "내역": "CMS 사용료(자동결제)",
            "지출대상자": "효성CMS",
            "*과목": "후원금모금경비",
            "*수입지출처구분": "",
            "*증빙서첨부": "",
            "*지출방법": "",
            "금액없어도포함": False
        },
        {
            "은행입력키워드": "카카오페이정산",
            "내역": "수수료",
            "지출대상자": "카카오페이",
            "*과목": "후원금모금경비",
            "*수입지출처구분": "",
            "*증빙서첨부": "",
            "*지출방법": "",
            "금액없어도포함": True
        }
    ],
    "fallback": {
        "*과목": "후원금모금경비",
        "*내역": "",
        "*지출대상자": "NH농협은행",
        "생년월일(사업자번호)": "",
        "우편번호": "",
        "주 소": "",
        "상세주소": "",
        "직업(업종)": "",
        "전화번호": "",
        "*증빙서첨부": "",
        "*수입지출처구분": "",
        "*지출방법": ""
    }
}


def normalize_loaded_config(config_data):
    if not isinstance(config_data, dict):
        return DEFAULT_EXPENSE_CONFIG

    targets = config_data.get("targets", [])
    rules = config_data.get("rules", [])
    fallback = config_data.get("fallback", {})

    if not isinstance(targets, list):
        targets = []
    if not isinstance(rules, list):
        rules = []
    if not isinstance(fallback, dict):
        fallback = {}

    normalized_rules = []
    for rule in rules:
        if not isinstance(rule, dict):
            continue

        normalized_rules.append({
            "은행입력키워드": rule.get("은행입력키워드", ""),
            "내역": rule.get("내역", ""),
            "지출대상자": rule.get("지출대상자", ""),
            "*과목": rule.get("*과목", "후원금모금경비"),
            "*수입지출처구분": rule.get("*수입지출처구분", ""),
            "*증빙서첨부": rule.get("*증빙서첨부", ""),
            "*지출방법": rule.get("*지출방법", ""),
            "금액없어도포함": rule.get("금액없어도포함", False),
        })

    normalized_fallback = {
        "*과목": fallback.get("*과목", "후원금모금경비"),
        "*내역": fallback.get("*내역", ""),
        "*지출대상자": fallback.get("*지출대상자", "NH농협은행"),
        "생년월일(사업자번호)": fallback.get("생년월일(사업자번호)", ""),
        "우편번호": fallback.get("우편번호", ""),
        "주 소": fallback.get("주 소", ""),
        "상세주소": fallback.get("상세주소", ""),
        "직업(업종)": fallback.get("직업(업종)", ""),
        "전화번호": fallback.get("전화번호", ""),
        "*증빙서첨부": fallback.get("*증빙서첨부", ""),
        "*수입지출처구분": fallback.get("*수입지출처구분", ""),
        "*지출방법": fallback.get("*지출방법", "")
    }

    return {
        "targets": targets,
        "rules": normalized_rules,
        "fallback": normalized_fallback
    }


def to_bool(v):
    if isinstance(v, bool):
        return v
    if str(v).strip().lower() in ["true", "1", "yes", "y"]:
        return True
    return False


if "expense_config_data" not in st.session_state:
    st.session_state["expense_config_data"] = DEFAULT_EXPENSE_CONFIG


st.subheader("1. 은행 거래내역 업로드")

uploaded_file = st.file_uploader("은행 거래내역 업로드", type=["xlsx", "xls"])

if uploaded_file:
    bank_file_path = os.path.join(INPUT_DIR, uploaded_file.name)

    with open(bank_file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.success("업로드 완료")

    col1, col2 = st.columns(2)

    with col1:
        if st.button("수입등록파일 생성하기"):
            output_path = os.path.join(
                OUTPUT_DIR,
                "후원회_프로그램_수입_일괄등록_결과.xlsx"
            )

            try:
                df_income, saved_income_path = run_donation_pipeline(
                    bank_file_path=bank_file_path,
                    template_file_path=INCOME_TEMPLATE_PATH,
                    output_file_path=output_path
                )

                if df_income.empty:
                    st.warning("수입으로 처리된 데이터가 없습니다.")
                else:
                    st.success(f"수입등록파일 생성 완료 ({len(df_income)}건)")
                    with open(saved_income_path, "rb") as f:
                        st.download_button(
                            label="수입 파일 다운로드",
                            data=f,
                            file_name="후원회_프로그램_수입_일괄등록_결과.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_income_file"
                        )
            except Exception as e:
                st.error(f"수입 파일 생성 오류: {e}")

    with col2:
        if st.button("지출등록파일 생성하기"):
            output_path = os.path.join(
                OUTPUT_DIR,
                "후원회_프로그램_지출_일괄등록_결과.xlsx"
            )

            try:
                current_config = {
                    "targets": st.session_state.get(
                        "edited_targets_records",
                        st.session_state["expense_config_data"]["targets"]
                    ),
                    "rules": st.session_state.get(
                        "edited_rules_records",
                        st.session_state["expense_config_data"]["rules"]
                    ),
                    "fallback": st.session_state.get(
                        "edited_fallback_record",
                        st.session_state["expense_config_data"]["fallback"]
                    )
                }

                df_expense, saved_expense_path = run_expense_pipeline(
                    bank_file_path=bank_file_path,
                    template_file_path=EXPENSE_TEMPLATE_PATH,
                    output_file_path=output_path,
                    config_data=current_config
                )

                if df_expense.empty:
                    st.warning("지출로 들어갈 데이터가 없습니다.")
                else:
                    st.success(f"지출등록파일 생성 완료 ({len(df_expense)}건)")
                    show_cols = [
                        col for col in [
                            "거래일자", "출금금액(원)", "거래기록사항",
                            "*내역", "*지출대상자", "*과목",
                            "*수입지출처구분", "*증빙서첨부", "*지출방법", "*금액"
                        ]
                        if col in df_expense.columns
                    ]
                    st.dataframe(df_expense[show_cols], use_container_width=True)

                    with open(saved_expense_path, "rb") as f:
                        st.download_button(
                            label="지출 파일 다운로드",
                            data=f,
                            file_name="후원회_프로그램_지출_일괄등록_결과.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_expense_file"
                        )

            except Exception as e:
                st.error(f"지출 파일 생성 오류: {e}")

st.divider()

st.subheader("2. 지출 설정 관리")

config_top_col1, config_top_col2 = st.columns(2)

with config_top_col1:
    uploaded_json = st.file_uploader(
        "지출 설정 JSON 업로드",
        type=["json"],
        key="expense_json_uploader"
    )

    if uploaded_json is not None:
        try:
            uploaded_config = json.load(uploaded_json)
            uploaded_config = normalize_loaded_config(uploaded_config)
            st.session_state["expense_config_data"] = uploaded_config
            st.success("JSON 업로드 및 적용 완료")
        except Exception as e:
            st.error(f"JSON 업로드 오류: {e}")

with config_top_col2:
    if st.button("기본 설정으로 초기화"):
        st.session_state["expense_config_data"] = DEFAULT_EXPENSE_CONFIG
        st.success("기본 설정으로 초기화 완료")

config_data = st.session_state["expense_config_data"]

targets_df = pd.DataFrame(config_data.get("targets", []))
rules_df = pd.DataFrame(config_data.get("rules", []))
fallback_df = pd.DataFrame([config_data.get("fallback", DEFAULT_EXPENSE_CONFIG["fallback"])])

if "금액없어도포함" not in rules_df.columns:
    rules_df["금액없어도포함"] = False

rules_df["금액없어도포함"] = rules_df["금액없어도포함"].fillna(False)
rules_df["금액없어도포함"] = rules_df["금액없어도포함"].apply(to_bool)

tab1, tab2, tab3 = st.tabs(["지출대상자 정보", "거래기록사항 규칙", "기본 fallback 설정"])

with tab1:
    st.caption("지출대상자별 고정값을 입력하거나 수정하세요.")
    edited_targets_df = st.data_editor(
        targets_df,
        num_rows="dynamic",
        use_container_width=True,
        key="expense_targets_editor"
    )
    st.session_state["edited_targets_records"] = edited_targets_df.fillna("").to_dict(orient="records")

with tab2:
    st.caption("은행 내역에 어떻게 적히면 어떤 값으로 넣을지 규칙을 입력하거나 수정하세요.")
    st.caption("체크하면 출금금액이 없어도 키워드만 맞으면 지출 파일에 포함됩니다.")

    edited_rules_df = st.data_editor(
        rules_df,
        num_rows="dynamic",
        use_container_width=True,
        key="expense_rules_editor",
        column_config={
            "은행입력키워드": st.column_config.TextColumn("은행입력키워드"),
            "내역": st.column_config.TextColumn("내역"),
            "지출대상자": st.column_config.TextColumn("지출대상자"),
            "*과목": st.column_config.TextColumn("*과목"),
            "*수입지출처구분": st.column_config.TextColumn("*수입지출처구분"),
            "*증빙서첨부": st.column_config.TextColumn("*증빙서첨부"),
            "*지출방법": st.column_config.TextColumn("*지출방법"),
            "금액없어도포함": st.column_config.CheckboxColumn(
                "금액 없을 때도 포함",
                default=False,
                help="체크하면 출금금액이 없어도 키워드가 맞으면 지출 파일에 포함합니다."
            ),
        }
    )

    rules_to_save = edited_rules_df.copy()

    if "금액없어도포함" not in rules_to_save.columns:
        rules_to_save["금액없어도포함"] = False

    rules_to_save["금액없어도포함"] = rules_to_save["금액없어도포함"].fillna(False)
    rules_to_save["금액없어도포함"] = rules_to_save["금액없어도포함"].apply(to_bool)

    for col in rules_to_save.columns:
        if col != "금액없어도포함":
            rules_to_save[col] = rules_to_save[col].fillna("")

    st.session_state["edited_rules_records"] = rules_to_save.to_dict(orient="records")

with tab3:
    st.caption("규칙에 맞지 않지만 출금금액이 있는 경우 자동으로 넣을 기본값을 설정하세요.")

    edited_fallback_df = st.data_editor(
        fallback_df,
        num_rows="fixed",
        use_container_width=True,
        key="expense_fallback_editor",
        column_config={
            "*과목": st.column_config.TextColumn("*과목"),
            "*내역": st.column_config.TextColumn("*내역"),
            "*지출대상자": st.column_config.TextColumn("*지출대상자"),
            "생년월일(사업자번호)": st.column_config.TextColumn("생년월일(사업자번호)"),
            "우편번호": st.column_config.TextColumn("우편번호"),
            "주 소": st.column_config.TextColumn("주 소"),
            "상세주소": st.column_config.TextColumn("상세주소"),
            "직업(업종)": st.column_config.TextColumn("직업(업종)"),
            "전화번호": st.column_config.TextColumn("전화번호"),
            "*증빙서첨부": st.column_config.TextColumn("*증빙서첨부"),
            "*수입지출처구분": st.column_config.TextColumn("*수입지출처구분"),
            "*지출방법": st.column_config.TextColumn("*지출방법"),
        }
    )

    fallback_to_save = edited_fallback_df.fillna("").to_dict(orient="records")[0]
    st.session_state["edited_fallback_record"] = fallback_to_save

col3, col4 = st.columns(2)

with col3:
    if st.button("현재 설정 적용"):
        st.session_state["expense_config_data"] = {
            "targets": edited_targets_df.fillna("").to_dict(orient="records"),
            "rules": rules_to_save.to_dict(orient="records"),
            "fallback": fallback_to_save
        }
        st.success("현재 화면의 설정이 이번 세션에 적용되었습니다.")

with col4:
    download_json_text = json.dumps(
        {
            "targets": edited_targets_df.fillna("").to_dict(orient="records"),
            "rules": rules_to_save.to_dict(orient="records"),
            "fallback": fallback_to_save
        },
        ensure_ascii=False,
        indent=2
    )

    st.download_button(
        label="지출 설정 JSON 다운로드",
        data=download_json_text,
        file_name="expense_config.json",
        mime="application/json",
        key="download_expense_config_json"
    )