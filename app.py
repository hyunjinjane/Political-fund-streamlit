import os
import json
import tempfile
from io import BytesIO
from typing import Any

import streamlit as st
from openpyxl import Workbook

from pipeline import run_pipeline


# =========================
# 기준파일 고정 경로
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FIXED_PATH = os.path.join(BASE_DIR, "data", "input", "정치자금_지출.xlsx")


# =========================
# 컬럼 정의
# =========================
DESC_COLS = ["keyword", "value", "job"]

PARTY_COLS = [
    "내역",
    "지출대상자",
    "생년월일(사업자번호)",
    "주소",
    "직업(업종)",
    "전화번호",
    "수입지출처구분",
]


# =========================
# 기본 규칙
# =========================
DEFAULT_DESC_RULES = [
    {"keyword": "주유소", "value": "수행주유비", "job": "주유"},
    {"keyword": "택시", "value": "수행택시비", "job": "택시"},
    {"keyword": "입력", "value": "입력", "job": "입력"},
]

DEFAULT_PARTY_RULES = [
    {
        "내역": "입력",
        "지출대상자": "입력",
        "생년월일(사업자번호)": "입력",
        "주소": "입력",
        "직업(업종)": "입력",
        "전화번호": "입력",
        "수입지출처구분": "입력",
    }
]


# =========================
# 유틸
# =========================
def normalize_table_rows(rows: list[dict], columns: list[str]) -> list[dict]:
    """
    - 컬럼 누락 시 빈 문자열
    - None -> ""
    - 숫자 들어오면 문자열로 저장(사업자번호/전화번호)
    - 완전 빈 행 제거
    """
    norm: list[dict] = []
    for r in rows or []:
        row: dict[str, str] = {}
        is_all_empty = True
        for c in columns:
            v = r.get(c, "")
            if v is None:
                v = ""
            if isinstance(v, (int, float)):
                v = str(int(v)) if float(v).is_integer() else str(v)
            v = str(v)
            row[c] = v
            if v.strip() != "":
                is_all_empty = False
        if not is_all_empty:
            norm.append(row)
    return norm


def ensure_party_rules_has_desc(rows: list[dict]) -> list[dict]:
    for r in rows or []:
        if "내역" not in r:
            r["내역"] = ""
    return rows


def safe_load_rules_json(uploaded_file) -> dict:
    raw = uploaded_file.getvalue()
    try:
        data = json.loads(raw.decode("utf-8"))
    except Exception as e:
        raise ValueError(f"JSON 파싱 실패: {e}")

    if not isinstance(data, dict):
        raise ValueError("rules.json 최상위는 dict여야 합니다.")
    if "desc_rules" not in data or "party_rules" not in data:
        raise ValueError("rules.json에 desc_rules, party_rules 키가 필요합니다.")
    if not isinstance(data["desc_rules"], list) or not isinstance(data["party_rules"], list):
        raise ValueError("desc_rules와 party_rules는 list여야 합니다.")
    return data


def build_rules_json_bytes(desc_rules: list[dict], party_rules: list[dict]) -> bytes:
    data = {"version": 1, "desc_rules": desc_rules, "party_rules": party_rules}
    return json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")


def build_no_match_excel(no_match: list[tuple[str, str]]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "PDF_매칭실패"
    ws.append(["PDF파일명", "실패사유"])
    for name, reason in no_match:
        ws.append([name, reason])
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


def xls_to_xlsx_bytes(xls_bytes: bytes) -> bytes:
    """
    .xls 바이너리를 .xlsx 바이너리로 변환
    - pandas + xlrd 필요
    """
    try:
        import pandas as pd  # type: ignore
    except Exception:
        raise RuntimeError("xls 변환을 위해 pandas가 필요합니다. requirements.txt에 pandas를 추가하세요.")

    try:
        df = pd.read_excel(BytesIO(xls_bytes), engine="xlrd")
    except Exception as e:
        raise RuntimeError(
            "xls 파일을 읽을 수 없습니다. (xlrd 필요)\n"
            "requirements.txt에 xlrd==2.0.1 을 추가하고 다시 배포/설치하세요.\n"
            f"원인: {e}"
        )

    wb = Workbook()
    ws = wb.active
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append(list(row))

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def init_session_state() -> None:
    ss = st.session_state

    ss.setdefault("desc_rules", DEFAULT_DESC_RULES)
    ss.setdefault("party_rules", DEFAULT_PARTY_RULES)

    ss.setdefault("desc_rules_draft", ss["desc_rules"])
    ss.setdefault("party_rules_draft", ss["party_rules"])

    ss.setdefault("pending_rules", None)

    ss.setdefault("rules_download_bytes", None)
    ss.setdefault("rules_download_version", 0)


def apply_rules(desc_rows: list[dict], party_rows: list[dict]) -> None:
    ss = st.session_state
    desc_clean = normalize_table_rows(desc_rows, DESC_COLS)
    party_clean = normalize_table_rows(party_rows, PARTY_COLS)
    party_clean = ensure_party_rules_has_desc(party_clean)

    ss["desc_rules"] = desc_clean
    ss["party_rules"] = party_clean
    ss["desc_rules_draft"] = desc_clean
    ss["party_rules_draft"] = party_clean

    ss["rules_download_bytes"] = build_rules_json_bytes(desc_clean, party_clean)
    ss["rules_download_version"] += 1


def reset_rules() -> None:
    ss = st.session_state
    ss["desc_rules"] = DEFAULT_DESC_RULES
    ss["party_rules"] = DEFAULT_PARTY_RULES
    ss["desc_rules_draft"] = DEFAULT_DESC_RULES
    ss["party_rules_draft"] = DEFAULT_PARTY_RULES
    ss["pending_rules"] = None
    ss["rules_download_bytes"] = None
    ss["rules_download_version"] += 1


# =========================
# UI 기본
# =========================
st.set_page_config(page_title="정치자금 지출 정리", layout="wide")

st.markdown(
    """
    <style>
      .block-container{
        max-width: 1900px;
        padding-top: 1.5rem;
        padding-bottom: 2.0rem;
        padding-left: 2.0rem;
        padding-right: 2.0rem;
      }
      .hint{
        color: rgba(49, 51, 63, 0.65);
        font-size: 0.92rem;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

init_session_state()

st.title("정치자금 지출 정리 자동화")
st.caption("은행내역(xls/xlsx) + 매출전표(PDF)를 기준파일 형식으로 자동 정리합니다.")

if not os.path.exists(TEMPLATE_FIXED_PATH):
    st.error(
        "고정 기준파일을 찾지 못했습니다.\n\n"
        f"- 경로: {TEMPLATE_FIXED_PATH}\n\n"
        "해결: data/input/정치자금_지출.xlsx 로 기준파일을 복사해 주세요."
    )
    st.stop()

st.success("기준파일은 고정 템플릿을 사용합니다: data/input/정치자금_지출.xlsx")
st.divider()


# =========================
# 1) 파일 업로드
# =========================
st.subheader("1) 파일 업로드")

up_c1, up_c2 = st.columns([1.1, 1.3], gap="large")
with up_c1:
    bank_file = st.file_uploader("은행내역 업로드 (xls/xlsx)", type=["xls", "xlsx"])
with up_c2:
    pdf_files = st.file_uploader("매출전표 PDF 업로드 (여러 개 가능)", type=["pdf"], accept_multiple_files=True)
    st.markdown('<div class="hint">PDF가 없으면 주소 규칙으로 채우는 방식만 적용됩니다.</div>', unsafe_allow_html=True)

st.divider()


# =========================
# 2) 기본 설정  (✅ 고정값 가로 배치)
# =========================
st.subheader("2) 기본 설정")

a1, a2 = st.columns([1, 1], gap="large")
with a1:
    fixed_account = st.text_input("*계정(고정값)", value="후원회기부금")
with a2:
    fixed_subject = st.text_input("*과목(고정값)", value="선거비용의 정치자금")

# ✅ 체크박스 제거: 덮어쓰기 방지는 True로 고정
skip_overwrite = True
st.markdown('<div class="hint">※ 기존 값이 이미 있으면 덮어쓰지 않도록 고정되어 있습니다.</div>', unsafe_allow_html=True)

st.divider()


# =========================
# 3) 규칙 관리
# =========================
st.subheader("3) 규칙 관리")
st.caption("서버 저장 없음: rules.json은 다운로드해서 보관하고, 필요할 때 업로드해서 사용하세요.")

uploaded_rules = st.file_uploader("rules.json 불러오기 (필요할 때만)", type=["json"], key="rules_json_uploader")
if uploaded_rules is not None:
    try:
        data = safe_load_rules_json(uploaded_rules)
        desc_loaded = normalize_table_rows(data["desc_rules"], DESC_COLS)
        party_loaded = normalize_table_rows(data["party_rules"], PARTY_COLS)
        party_loaded = ensure_party_rules_has_desc(party_loaded)

        st.session_state["pending_rules"] = {"desc_rules": desc_loaded, "party_rules": party_loaded}
        st.success("rules.json을 불러왔습니다. 아래 4)에서 '변경사항 적용'을 누르면 표에 반영됩니다.")
    except Exception as e:
        st.session_state["pending_rules"] = None
        st.error(str(e))

st.divider()


# =========================
# 4) 규칙 편집  (✅ 다운로드 버튼 복구)
# =========================
st.subheader("4) 규칙 편집")
st.caption("표 수정 → '변경사항 적용' → 가운데 rules.json 다운로드로 저장하세요.")

tab1, tab2 = st.tabs(["내역 규칙", "주소 규칙"])

with tab1:
    st.markdown(
        '<div class="hint">keyword가 <b>지출대상자에 포함</b>되면 value(내역) + job(직업)을 채웁니다. 위에서부터 <b>첫 매칭</b>만 적용(포함).</div>',
        unsafe_allow_html=True,
    )
    edited_desc = st.data_editor(
        st.session_state["desc_rules_draft"],
        num_rows="dynamic",
        use_container_width=True,
        column_order=DESC_COLS,
        key="desc_rules_editor",
    )

with tab2:
    st.markdown(
        '<div class="hint">※ <b>지출대상자 완전 동일</b>한 경우에만 아래 값이 들어갑니다. (띄어쓰기 차이는 무시)</div>',
        unsafe_allow_html=True,
    )
    edited_party = st.data_editor(
        st.session_state["party_rules_draft"],
        num_rows="dynamic",
        use_container_width=True,
        column_order=PARTY_COLS,
        key="party_rules_editor",
    )

# ✅ 버튼 3개 한 줄(왼:적용 / 중:다운로드 / 우:초기화) + 가운데 자리 고정
btn_apply, btn_download, btn_reset = st.columns([1.2, 1.2, 0.9], gap="large")

with btn_apply:
    apply_clicked = st.button("✅ 변경사항 적용", type="primary", use_container_width=True)

with btn_download:
    if st.session_state["rules_download_bytes"] is None:
        st.button("⬇️ rules.json 다운로드", disabled=True, use_container_width=True)
        st.markdown('<div class="hint">먼저 왼쪽에서 “변경사항 적용”을 눌러주세요.</div>', unsafe_allow_html=True)
    else:
        st.download_button(
            label="⬇️ rules.json 다운로드",
            data=st.session_state["rules_download_bytes"],
            file_name="rules.json",
            mime="application/json",
            use_container_width=True,
            key=f"rules_dl_{st.session_state['rules_download_version']}",
        )
        st.markdown('<div class="hint">이 파일을 저장해두면 다음에 그대로 불러올 수 있어요.</div>', unsafe_allow_html=True)

with btn_reset:
    reset_clicked = st.button("♻️ 초기화", use_container_width=True)

if apply_clicked:
    if st.session_state.get("pending_rules") is not None:
        pending = st.session_state["pending_rules"]
        apply_rules(pending["desc_rules"], pending["party_rules"])
        st.session_state["pending_rules"] = None
    else:
        apply_rules(edited_desc, edited_party)

    st.success("적용 완료! 이제 가운데에서 rules.json을 다운로드하세요.")
    st.rerun()

if reset_clicked:
    reset_rules()
    st.success("기본 규칙으로 초기화했습니다.")
    st.rerun()

st.divider()


# =========================
# 5) 실행
# =========================
st.subheader("5) 실행")
run_btn = st.button("정리 실행", type="primary", use_container_width=True)

if run_btn:
    if not bank_file:
        st.error("은행내역 파일은 필수입니다.")
        st.stop()

    # 실행 직전에 clean
    desc_rules_clean = normalize_table_rows(st.session_state["desc_rules"], DESC_COLS)
    party_rules_clean = normalize_table_rows(st.session_state["party_rules"], PARTY_COLS)
    party_rules_clean = ensure_party_rules_has_desc(party_rules_clean)

    with st.spinner("처리 중..."):
        with tempfile.TemporaryDirectory(prefix="politics_") as tmp:
            # 기준 템플릿 복사
            template_path = os.path.join(tmp, "template.xlsx")
            with open(TEMPLATE_FIXED_PATH, "rb") as src, open(template_path, "wb") as dst:
                dst.write(src.read())

            # 은행파일 저장(여기서 xls면 xlsx로 변환)
            bank_upload_bytes = bank_file.getvalue()
            bank_name = (bank_file.name or "").lower()
            bank_path = os.path.join(tmp, "bank.xlsx")

            if bank_name.endswith(".xls"):
                try:
                    converted = xls_to_xlsx_bytes(bank_upload_bytes)
                    with open(bank_path, "wb") as f:
                        f.write(converted)
                except Exception as e:
                    st.error(str(e))
                    st.stop()
            else:
                with open(bank_path, "wb") as f:
                    f.write(bank_upload_bytes)

            # pdf 저장
            pdf_dir = os.path.join(tmp, "pdfs")
            os.makedirs(pdf_dir, exist_ok=True)
            for pf in pdf_files or []:
                with open(os.path.join(pdf_dir, pf.name), "wb") as f:
                    f.write(pf.getbuffer())

            output_path = os.path.join(tmp, "정리결과.xlsx")

            result: dict[str, Any] = run_pipeline(
                template_path=template_path,
                bank_path=bank_path,
                pdf_dir=pdf_dir,
                output_path=output_path,
                fixed_account=fixed_account,
                fixed_subject=fixed_subject,
                desc_rules=desc_rules_clean,
                party_rules=party_rules_clean,
                skip_if_already_filled=skip_overwrite,  # ✅ True 고정
            )

            st.success("완료!")

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("은행추가", result.get("bank_rows_added", 0))
            m2.metric("PDF보강", result.get("pdf_updated_rows", 0))
            m3.metric("주소규칙 보강(셀)", result.get("partyinfo_filled_cells", 0))
            m4.metric("PDF 실패", len(result.get("no_match", [])))

            with st.expander("실행 로그 보기"):
                for line in result.get("logs", []):
                    st.write("• " + str(line))

            with open(output_path, "rb") as f:
                st.download_button(
                    label="정리결과.xlsx 다운로드",
                    data=f,
                    file_name="정리결과.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            if result.get("no_match"):
                no_match_xlsx = build_no_match_excel(result["no_match"])
                st.download_button(
                    label="PDF 매칭 실패목록.xlsx 다운로드",
                    data=no_match_xlsx,
                    file_name="PDF_매칭실패목록.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
