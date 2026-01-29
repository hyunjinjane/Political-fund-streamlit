import os
import json
import tempfile
from io import BytesIO

import streamlit as st
from openpyxl import Workbook

from pipeline import run_pipeline


# =========================
# 기준파일 고정 경로
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FIXED_PATH = os.path.join(BASE_DIR, "data", "input", "정치자금_지출.xlsx")


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
    norm = []
    for r in rows or []:
        row = {}
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


# =========================
# 컬럼 정의
# =========================
DESC_COLS = ["keyword", "value", "job"]

# 주소 규칙에 내역 포함(이전 요청 반영)
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
    {"keyword": "주유소", "value": "수행주유비", "job": ""},
    {"keyword": "택시", "value": "수행택시비", "job": ""},
    {"keyword": "입력", "value": "입력", "job": ""},
]

DEFAULT_PARTY_RULES = [
    {
        "지출대상자": "상호명",
        "생년월일(사업자번호)": "사업자번호",
        "주소": "입력",
        "직업(업종)": "입력",
        "전화번호": "입력",
        "수입지출처구분": "입력",
        "내역": "입력",
    }
]


# =========================
# UI
# =========================
st.set_page_config(page_title="정치자금 지출 정리", layout="centered")
st.title("정치자금 지출 정리 자동화")
st.caption("은행내역(xlsx) + 매출전표(PDF)를 기준파일 형식으로 자동 정리합니다.")

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
# 세션 초기화
# =========================
if "desc_rules" not in st.session_state:
    st.session_state["desc_rules"] = DEFAULT_DESC_RULES
if "party_rules" not in st.session_state:
    st.session_state["party_rules"] = DEFAULT_PARTY_RULES

# 편집용 draft (IME 안정 위해 form에서만 확정)
if "desc_rules_draft" not in st.session_state:
    st.session_state["desc_rules_draft"] = st.session_state["desc_rules"]
if "party_rules_draft" not in st.session_state:
    st.session_state["party_rules_draft"] = st.session_state["party_rules"]

# 업로드한 rules.json 임시 보관(적용 전)
if "pending_rules" not in st.session_state:
    st.session_state["pending_rules"] = None
if "pending_loaded_msg" not in st.session_state:
    st.session_state["pending_loaded_msg"] = False

# 다운로드 준비 상태
if "rules_download_ready" not in st.session_state:
    st.session_state["rules_download_ready"] = None
if "rules_download_version" not in st.session_state:
    st.session_state["rules_download_version"] = 0


# -------------------------
# 1) 파일 업로드
# -------------------------
st.subheader("1) 파일 업로드")
bank_file = st.file_uploader("은행내역 업로드 (xlsx)", type=["xlsx"])
pdf_files = st.file_uploader("매출전표 PDF 업로드 (여러 개 가능)", type=["pdf"], accept_multiple_files=True)

st.divider()

# -------------------------
# 2) 기본 설정
# -------------------------
st.subheader("2) 기본 설정")
cA, cB = st.columns(2)
with cA:
    fixed_account = st.text_input("*계정(고정값)", value="후원회기부금")
with cB:
    fixed_subject = st.text_input("*과목(고정값)", value="선거비용의 정치자금")
skip_overwrite = st.checkbox("주소/사업자번호/전화번호/직업/수입지출처구분/내역이 이미 있으면 덮어쓰지 않기", value=True)

st.divider()

# -------------------------
# 3) 규칙 관리
# -------------------------
st.subheader("3) 규칙 관리")
st.caption("서버 저장 없음: rules.json을 다운로드해 보관하고, 필요할 때 업로드해서 사용하세요.")

col1, col2, col3 = st.columns([1.2, 1.0, 1.0], gap="large")

with col1:
    st.markdown("#### 📥 불러오기")
    uploaded_rules = st.file_uploader(
        "rules.json 업로드",
        type=["json"],
        key="rules_json_uploader",
        label_visibility="collapsed",
    )
    if uploaded_rules is not None:
        try:
            data = safe_load_rules_json(uploaded_rules)
            desc_loaded = normalize_table_rows(data["desc_rules"], DESC_COLS)
            party_loaded = normalize_table_rows(data["party_rules"], PARTY_COLS)
            # 구버전 호환(내역 컬럼)
            for r in party_loaded:
                if "내역" not in r:
                    r["내역"] = ""

            st.session_state["pending_rules"] = {
                "desc_rules": desc_loaded,
                "party_rules": party_loaded,
            }
            st.session_state["pending_loaded_msg"] = True
        except Exception as e:
            st.session_state["pending_rules"] = None
            st.session_state["pending_loaded_msg"] = False
            st.error(str(e))

with col2:
    st.markdown("#### 🧾 내보내기")
    st.caption("아래 '규칙 편집'에서 편집 → 저장(다운로드 준비) 후 다운로드하세요.")

    if st.session_state["rules_download_ready"] is None:
        st.button("rules.json 다운로드", disabled=True, use_container_width=True)
    else:
        st.download_button(
            label="rules.json 다운로드",
            data=st.session_state["rules_download_ready"],
            file_name="rules.json",
            mime="application/json",
            use_container_width=True,
            key=f"rules_dl_{st.session_state['rules_download_version']}",
        )

with col3:
    st.markdown("#### ♻️ 초기화")
    if st.button("기본 규칙으로 초기화", use_container_width=True):
        st.session_state["desc_rules"] = DEFAULT_DESC_RULES
        st.session_state["party_rules"] = DEFAULT_PARTY_RULES
        st.session_state["desc_rules_draft"] = DEFAULT_DESC_RULES
        st.session_state["party_rules_draft"] = DEFAULT_PARTY_RULES
        st.session_state["pending_rules"] = None
        st.session_state["pending_loaded_msg"] = False
        st.session_state["rules_download_ready"] = None
        st.success("기본 규칙으로 초기화했습니다.")
        st.rerun()

if st.session_state["pending_rules"] is not None:
    if st.session_state["pending_loaded_msg"]:
        st.success("rules.json을 불러왔습니다. 아래 버튼을 눌러 규칙 편집 표에 반영하세요.")
        st.session_state["pending_loaded_msg"] = False

    st.warning("아직 표에 반영되지 않았습니다.", icon="⚠️")

    if st.button("✅ 규칙 편집 표에 반영하기", type="primary", use_container_width=True):
        pending = st.session_state["pending_rules"]
        st.session_state["desc_rules"] = pending["desc_rules"]
        st.session_state["party_rules"] = pending["party_rules"]
        st.session_state["desc_rules_draft"] = pending["desc_rules"]
        st.session_state["party_rules_draft"] = pending["party_rules"]
        st.session_state["pending_rules"] = None
        st.session_state["rules_download_ready"] = None
        st.success("표에 반영했습니다! (아래에서 편집 후 저장(다운로드 준비) 해주세요)")
        st.rerun()

st.divider()

# -------------------------
# 4) 규칙 편집 (저장 버튼 제거 → "다운로드 준비" 버튼 1개로 통합)
# -------------------------
st.subheader("4) 규칙 편집")
st.caption("표를 편집한 뒤, 아래의 'rules.json 저장(다운로드 준비)' 버튼을 누르면 편집 내용이 확정됩니다.")

tab1, tab2 = st.tabs(["내역 규칙", "주소 규칙"])

with st.form("rules_edit_form", clear_on_submit=False):
    with tab1:
        st.caption("keyword가 *지출대상자에 포함되면 value(내역) + job(직업)을 채웁니다. 위에서부터 첫 매칭만 적용(포함).")
        desc_draft = st.data_editor(
            st.session_state["desc_rules_draft"],
            num_rows="dynamic",
            use_container_width=True,
            column_order=DESC_COLS,
            key="desc_rules_editor_form",
        )

    with tab2:
        st.caption("※ '*지출대상자'가 '완전히 동일'한 경우에만 아래 값이 들어갑니다. (띄어쓰기 차이는 무시)")
        party_draft = st.data_editor(
            st.session_state["party_rules_draft"],
            num_rows="dynamic",
            use_container_width=True,
            column_order=PARTY_COLS,
            key="party_rules_editor_form",
        )

    # ✅ 저장 버튼을 하나로 통합(다운로드 준비)
    prep = st.form_submit_button("💾 rules.json 저장(다운로드 준비)", use_container_width=True)

    if prep:
        # 1) draft를 세션에 확정
        st.session_state["desc_rules_draft"] = desc_draft
        st.session_state["party_rules_draft"] = party_draft
        st.session_state["desc_rules"] = desc_draft
        st.session_state["party_rules"] = party_draft

        # 2) clean 후 rules.json bytes 준비
        desc_clean = normalize_table_rows(st.session_state["desc_rules"], DESC_COLS)
        party_clean = normalize_table_rows(st.session_state["party_rules"], PARTY_COLS)

        # 구버전 호환 보정
        for r in party_clean:
            if "내역" not in r:
                r["내역"] = ""

        st.session_state["rules_download_ready"] = build_rules_json_bytes(desc_clean, party_clean)
        st.session_state["rules_download_version"] += 1

        st.success("다운로드 준비 완료! 위 '내보내기'에서 rules.json 다운로드 버튼을 눌러 저장하세요.")

st.divider()

# -------------------------
# 5) 실행
# -------------------------
st.subheader("5) 실행")
run_btn = st.button("정리 실행", type="primary", use_container_width=True)

if run_btn:
    if not bank_file:
        st.error("은행내역 파일은 필수입니다.")
        st.stop()

    # 실행 직전에 clean
    desc_rules_clean = normalize_table_rows(st.session_state["desc_rules"], DESC_COLS)
    party_rules_clean = normalize_table_rows(st.session_state["party_rules"], PARTY_COLS)

    with st.spinner("처리 중..."):
        with tempfile.TemporaryDirectory(prefix="politics_") as tmp:
            template_path = os.path.join(tmp, "template.xlsx")
            with open(TEMPLATE_FIXED_PATH, "rb") as src, open(template_path, "wb") as dst:
                dst.write(src.read())

            bank_path = os.path.join(tmp, "bank.xlsx")
            pdf_dir = os.path.join(tmp, "pdfs")
            os.makedirs(pdf_dir, exist_ok=True)
            output_path = os.path.join(tmp, "정리결과.xlsx")

            with open(bank_path, "wb") as f:
                f.write(bank_file.getbuffer())

            for pf in pdf_files or []:
                with open(os.path.join(pdf_dir, pf.name), "wb") as f:
                    f.write(pf.getbuffer())

            result = run_pipeline(
                template_path=template_path,
                bank_path=bank_path,
                pdf_dir=pdf_dir,
                output_path=output_path,
                fixed_account=fixed_account,
                fixed_subject=fixed_subject,
                desc_rules=desc_rules_clean,
                party_rules=party_rules_clean,
                skip_if_already_filled=skip_overwrite,
            )

            st.success("완료!")

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("은행추가", result.get("bank_rows_added", 0))
            m2.metric("PDF보강", result.get("pdf_updated_rows", 0))
            m3.metric("주소규칙 보강(셀)", result.get("partyinfo_filled_cells", 0))
            m4.metric("PDF 실패", len(result.get("no_match", [])))

            with st.expander("실행 로그 보기"):
                for line in result.get("logs", []):
                    st.write("• " + line)

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

