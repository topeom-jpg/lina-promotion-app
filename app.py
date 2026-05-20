from __future__ import annotations

from io import BytesIO
from typing import Iterable, Optional

import pandas as pd
import streamlit as st


APP_TITLE = "라이나 통합건강1 금시상 콜리스트 자동생성기"

TARGET_PRODUCT_KEYWORDS = [
    "새로담는건강보험",
    "새로담는간편건강보험",
    "새로담는건강보험플러스",
]

TIERS = [
    {"tier": 50_000, "tier_name": "5만원 이상", "prize": 600_000, "prize_name": "60만원"},
    {"tier": 100_000, "tier_name": "10만원 이상", "prize": 1_400_000, "prize_name": "140만원"},
    {"tier": 200_000, "tier_name": "20만원 이상", "prize": 3_200_000, "prize_name": "320만원"},
    {"tier": 300_000, "tier_name": "30만원 이상", "prize": 5_400_000, "prize_name": "540만원"},
    {"tier": 500_000, "tier_name": "50만원 이상", "prize": 10_000_000, "prize_name": "1,000만원"},
]

REQUIRED_HEADERS = ["설계사", "상품명"]
RECOMMENDED_HEADERS = ["대리점명", "지점명", "계약상태", "납입상태", "보험료"]


# -------------------------------
# 공통 유틸
# -------------------------------

def get_secret(name: str, default: str = "") -> str:
    """Streamlit Cloud secrets가 없을 때도 로컬 실행이 깨지지 않도록 처리합니다."""
    try:
        return str(st.secrets.get(name, default))
    except Exception:
        return default


def gate_with_password() -> None:
    app_password = get_secret("APP_PASSWORD", "")
    if not app_password:
        st.sidebar.info("배포 전 .streamlit/secrets.toml 또는 Streamlit Secrets에 APP_PASSWORD를 설정하세요.")
        return

    if st.session_state.get("authenticated") is True:
        return

    st.title(APP_TITLE)
    st.caption("접속 비밀번호를 입력하면 업로드 화면이 열립니다.")
    user_pw = st.text_input("접속 비밀번호", type="password")
    if st.button("접속하기", use_container_width=True):
        if user_pw == app_password:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("비밀번호가 맞지 않습니다.")
    st.stop()


def money(value: object) -> str:
    if pd.isna(value):
        return "-"
    try:
        return f"{int(round(float(value))):,}원"
    except Exception:
        return str(value)


def normalize_colname(col: object) -> str:
    return str(col).strip().replace("\n", " ").replace("\r", " ")


def to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("원", "", regex=False)
        .str.replace(" ", "", regex=False),
        errors="coerce",
    ).fillna(0)


def find_header_row(raw: pd.DataFrame) -> int:
    """상단 제목행이 있어도 실제 헤더행을 자동 탐색합니다."""
    best_idx = 0
    best_score = -1
    header_words = set(REQUIRED_HEADERS + RECOMMENDED_HEADERS + ["계약번호", "계약일자", "피보험자명", "계약자명"])

    max_rows = min(len(raw), 30)
    for idx in range(max_rows):
        row_values = {normalize_colname(v) for v in raw.iloc[idx].tolist() if pd.notna(v)}
        score = len(row_values.intersection(header_words))
        if score > best_score:
            best_score = score
            best_idx = idx

    if best_score < len(REQUIRED_HEADERS):
        raise ValueError("헤더행을 찾지 못했습니다. '설계사', '상품명' 컬럼이 있는 엑셀인지 확인해주세요.")
    return best_idx


def read_excel_safely(uploaded_file, sheet_name: str) -> pd.DataFrame:
    raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, dtype=object)
    header_row = find_header_row(raw)
    headers = [normalize_colname(v) for v in raw.iloc[header_row].tolist()]
    df = raw.iloc[header_row + 1 :].copy()
    df.columns = headers
    df = df.dropna(how="all")
    df = df.loc[:, [c for c in df.columns if c and c != "nan"]]
    return df.reset_index(drop=True)


def validate_columns(df: pd.DataFrame) -> list[str]:
    missing = [c for c in REQUIRED_HEADERS if c not in df.columns]
    return missing


def contains_any(series: pd.Series, keywords: Iterable[str]) -> pd.Series:
    text = series.fillna("").astype(str)
    mask = pd.Series(False, index=series.index)
    for keyword in keywords:
        mask = mask | text.str.contains(keyword, regex=False, na=False)
    return mask


def current_tier_info(total_p: float) -> dict:
    current = {"tier": 0, "tier_name": "미달성", "prize": 0, "prize_name": "0원"}
    next_item: Optional[dict] = TIERS[0]

    for item in TIERS:
        if total_p >= item["tier"]:
            current = item
        else:
            next_item = item
            break
    else:
        next_item = None

    if next_item is None:
        shortage = 0
        next_tier = None
        next_tier_name = "최고구간 달성"
        next_prize = current["prize"]
        next_prize_name = current["prize_name"]
        increase = 0
        priority = "완료"
    else:
        shortage = max(0, int(next_item["tier"] - total_p))
        next_tier = next_item["tier"]
        next_tier_name = next_item["tier_name"]
        next_prize = next_item["prize"]
        next_prize_name = next_item["prize_name"]
        increase = int(next_prize - current["prize"])
        if shortage <= 20_000:
            priority = "S급"
        elif shortage <= 50_000:
            priority = "A급"
        else:
            priority = "B급"

    return {
        "현재구간": current["tier"],
        "현재구간명": current["tier_name"],
        "다음구간": next_tier,
        "다음구간명": next_tier_name,
        "부족P": shortage,
        "현재금시상": current["prize"],
        "현재금시상명": current["prize_name"],
        "다음금시상": next_prize,
        "다음금시상명": next_prize_name,
        "추가상승액": increase,
        "콜우선순위": priority,
    }


def make_call_message(row: pd.Series) -> str:
    current_p = int(row["현재 통합건강1 P"])
    priority = row["콜우선순위"]

    if priority == "완료":
        return f"현재 통합건강1 인정P {current_p:,}원으로 50만원 최고구간 달성입니다. 유지/철회 방어 체크가 우선입니다."

    shortage = int(row["부족P"])
    next_tier_name = row["다음구간명"]
    next_prize = row["다음금시상명"]
    increase = int(row["추가상승액"])

    if current_p == 0:
        return f"통합건강1 가동이 없습니다. {shortage:,}원 설계 시 {next_tier_name} 구간 진입, 13회차 금시상 {next_prize} 대상입니다."

    return (
        f"현재 통합건강1 인정P {current_p:,}원입니다. "
        f"{shortage:,}원만 추가하면 {next_tier_name} 구간 진입, "
        f"13회차 금시상 {next_prize} 대상입니다. 현재 대비 추가상승액은 {increase:,}원입니다."
    )


# -------------------------------
# 핵심 분석 로직
# -------------------------------

def analyze_promotion(df: pd.DataFrame, premium_col: str, filter_active: bool = True, filter_normal_payment: bool = True) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    work = df.copy()
    work.columns = [normalize_colname(c) for c in work.columns]

    for col in ["대리점명", "지점명", "계약상태", "납입상태"]:
        if col not in work.columns:
            work[col] = ""

    if premium_col not in work.columns:
        raise ValueError(f"보험료 기준 컬럼 '{premium_col}'을 찾을 수 없습니다.")

    work["상품명"] = work["상품명"].fillna("").astype(str)
    work["보험료_숫자"] = to_number(work[premium_col])
    work["대상상품여부"] = contains_any(work["상품명"], TARGET_PRODUCT_KEYWORDS)

    if "계약상태" in work.columns:
        work["계약상태"] = work["계약상태"].fillna("").astype(str).str.strip()
    if "납입상태" in work.columns:
        work["납입상태"] = work["납입상태"].fillna("").astype(str).str.strip()

    status_mask = True
    if filter_active and "계약상태" in work.columns:
        status_mask = work["계약상태"].eq("유지")
    payment_mask = True
    if filter_normal_payment and "납입상태" in work.columns:
        payment_mask = work["납입상태"].eq("정상")

    target = work[work["대상상품여부"] & status_mask & payment_mask].copy()
    target["인정보험료"] = target["보험료_숫자"].clip(upper=300_000)

    grouping_cols = ["대리점명", "지점명", "설계사"]
    summary = (
        target.groupby(grouping_cols, dropna=False)
        .agg(
            대상계약건수=("설계사", "size"),
            원보험료합계=("보험료_숫자", "sum"),
            **{"현재 통합건강1 P": ("인정보험료", "sum")},
        )
        .reset_index()
    )

    if summary.empty:
        summary = pd.DataFrame(columns=grouping_cols + ["대상계약건수", "원보험료합계", "현재 통합건강1 P"])

    tier_df = summary["현재 통합건강1 P"].apply(current_tier_info).apply(pd.Series)
    call_list = pd.concat([summary, tier_df], axis=1)
    call_list["콜멘트"] = call_list.apply(make_call_message, axis=1) if not call_list.empty else []

    priority_order = {"S급": 1, "A급": 2, "B급": 3, "완료": 4}
    call_list["정렬키"] = call_list["콜우선순위"].map(priority_order).fillna(9)
    call_list = call_list.sort_values(
        by=["정렬키", "부족P", "추가상승액", "현재 통합건강1 P"],
        ascending=[True, True, False, False],
    ).drop(columns=["정렬키"]).reset_index(drop=True)

    target_display_cols = [c for c in df.columns if c in [
        "NO", "대리점명", "지점명", "설계사", "계약번호", "계약일자", "계약상태", "상품명",
        "보험료", "계약자명", "피보험자명", "가입금액", "월환산보험료", "연환산보험료", "CMP",
        "납입기간", "납입주기", "납입상태",
    ]]
    for extra in ["보험료_숫자", "인정보험료"]:
        if extra not in target_display_cols:
            target_display_cols.append(extra)
    target = target[target_display_cols].copy()

    metrics = {
        "대상계약건수": int(len(target)),
        "대상설계사수": int(call_list["설계사"].nunique()) if "설계사" in call_list.columns else 0,
        "총인정보험료": int(target["인정보험료"].sum()) if "인정보험료" in target.columns else 0,
        "S급": int((call_list["콜우선순위"] == "S급").sum()) if "콜우선순위" in call_list.columns else 0,
        "A급": int((call_list["콜우선순위"] == "A급").sum()) if "콜우선순위" in call_list.columns else 0,
        "완료": int((call_list["콜우선순위"] == "완료").sum()) if "콜우선순위" in call_list.columns else 0,
    }
    return call_list, target, metrics


# -------------------------------
# 엑셀 생성 로직
# -------------------------------

def autosize_columns(worksheet, dataframe: pd.DataFrame, start_col: int = 0, max_width: int = 42) -> None:
    for idx, col in enumerate(dataframe.columns):
        values = dataframe[col].astype(str).replace("nan", "")
        width = max(len(str(col)), int(values.map(len).quantile(0.95)) if len(values) else 0) + 2
        worksheet.set_column(start_col + idx, start_col + idx, min(max(width, 10), max_width))


def make_excel_file(call_list: pd.DataFrame, target: pd.DataFrame, metrics: dict) -> bytes:
    output = BytesIO()

    call_cols = [
        "대리점명", "지점명", "설계사", "대상계약건수", "현재 통합건강1 P", "현재구간명",
        "다음구간명", "부족P", "현재금시상", "다음금시상", "추가상승액", "콜우선순위", "콜멘트",
    ]
    call_export = call_list[[c for c in call_cols if c in call_list.columns]].copy()

    tier_table = pd.DataFrame(TIERS)
    tier_table = tier_table.rename(columns={
        "tier": "달성구간",
        "tier_name": "구간명",
        "prize": "13회차 금시상/현금",
        "prize_name": "지급액표시",
    })

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        call_export.to_excel(writer, sheet_name="콜리스트_통합건강1", index=False, startrow=0)
        target.to_excel(writer, sheet_name="대상계약", index=False, startrow=0)
        tier_table.to_excel(writer, sheet_name="시책조건", index=False, startrow=3)

        workbook = writer.book
        fmt_title = workbook.add_format({"bold": True, "font_size": 18, "font_color": "#FFFFFF", "bg_color": "#17365D", "align": "center", "valign": "vcenter"})
        fmt_subtitle = workbook.add_format({"bold": True, "font_size": 12, "font_color": "#1F2937", "bg_color": "#EAF2F8"})
        fmt_header = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "border": 1, "align": "center", "valign": "vcenter"})
        fmt_body = workbook.add_format({"border": 1, "valign": "vcenter"})
        fmt_num = workbook.add_format({"num_format": "#,##0", "border": 1, "valign": "vcenter"})
        fmt_note = workbook.add_format({"text_wrap": True, "border": 1, "valign": "top"})
        fmt_metric_label = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        fmt_metric_value = workbook.add_format({"bold": True, "num_format": "#,##0", "border": 1, "align": "right"})

        # 대시보드
        dash = workbook.add_worksheet("대시보드")
        writer.sheets["대시보드"] = dash
        dash.merge_range("A1:F1", "2026년 5월 라이나생명 통합건강1 금시상 콜작업용", fmt_title)
        dash.merge_range("A2:F2", "새로담는건강보험 / 새로담는간편건강보험 / 새로담는건강보험플러스만 합산", fmt_subtitle)
        dash.write("A4", "분석 기준", fmt_header)
        dash.write("B4", "내용", fmt_header)
        dash.write("A5", "대상 상품군", fmt_body)
        dash.write("B5", "새로담는건강보험 / 새로담는간편건강보험 / 새로담는건강보험플러스", fmt_body)
        dash.write("A6", "적용 시책", fmt_body)
        dash.write("B6", "통합건강1 13회차 금시상 구간만 반영", fmt_body)
        dash.write("A7", "인정 기준", fmt_body)
        dash.write("B7", "계약상태=유지, 납입상태=정상, 건당 30만원 한도", fmt_body)
        dash.write("A8", "콜 등급", fmt_body)
        dash.write("B8", "S급: 부족P 2만원 이하 / A급: 5만원 이하 / B급: 그 외 / 완료: 50만원 이상", fmt_body)

        dash.write("D4", "핵심 KPI", fmt_header)
        dash.write("E4", "값", fmt_header)
        kpis = [
            ("대상 계약 건수", metrics.get("대상계약건수", 0)),
            ("대상 설계사 수", metrics.get("대상설계사수", 0)),
            ("총 인정 보험료", metrics.get("총인정보험료", 0)),
            ("S급 콜 대상", metrics.get("S급", 0)),
            ("A급 콜 대상", metrics.get("A급", 0)),
            ("완료자", metrics.get("완료", 0)),
        ]
        for r, (label, value) in enumerate(kpis, start=5):
            dash.write(r - 1, 3, label, fmt_metric_label)
            dash.write(r - 1, 4, value, fmt_metric_value)

        dash.write("A11", "우선 콜 TOP 10", fmt_header)
        top_cols = ["콜우선순위", "대리점명", "지점명", "설계사", "현재 통합건강1 P", "부족P", "다음구간명", "다음금시상명"]
        top = call_list[call_list["콜우선순위"].isin(["S급", "A급"])].head(10)
        if top.empty:
            top = call_list.head(10)
        top = top[[c for c in top_cols if c in top.columns]].copy()
        start_row = 11
        for col_num, col_name in enumerate(top.columns):
            dash.write(start_row, col_num, col_name, fmt_header)
        for row_num, row in enumerate(top.itertuples(index=False), start=start_row + 1):
            for col_num, value in enumerate(row):
                if isinstance(value, (int, float)) and not pd.isna(value):
                    dash.write_number(row_num, col_num, float(value), fmt_num)
                else:
                    dash.write(row_num, col_num, value, fmt_body)
        dash.set_column("A:A", 16)
        dash.set_column("B:C", 18)
        dash.set_column("D:E", 18)
        dash.set_column("F:F", 16)
        dash.set_column("G:G", 18)
        dash.set_row(0, 28)

        # 콜리스트 서식
        ws_call = writer.sheets["콜리스트_통합건강1"]
        ws_call.freeze_panes(1, 0)
        ws_call.autofilter(0, 0, max(len(call_export), 1), max(len(call_export.columns) - 1, 0))
        for col_idx, _ in enumerate(call_export.columns):
            ws_call.write(0, col_idx, call_export.columns[col_idx], fmt_header)
        autosize_columns(ws_call, call_export)
        number_cols = ["대상계약건수", "현재 통합건강1 P", "부족P", "현재금시상", "다음금시상", "추가상승액"]
        for col_name in number_cols:
            if col_name in call_export.columns:
                idx = call_export.columns.get_loc(col_name)
                ws_call.set_column(idx, idx, 15, fmt_num)
        if "콜멘트" in call_export.columns:
            idx = call_export.columns.get_loc("콜멘트")
            ws_call.set_column(idx, idx, 70, fmt_note)
        if "콜우선순위" in call_export.columns and len(call_export) > 0:
            priority_col = call_export.columns.get_loc("콜우선순위")
            ws_call.conditional_format(1, priority_col, len(call_export), priority_col, {
                "type": "text", "criteria": "containing", "value": "S급", "format": workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "bold": True, "border": 1})
            })
            ws_call.conditional_format(1, priority_col, len(call_export), priority_col, {
                "type": "text", "criteria": "containing", "value": "A급", "format": workbook.add_format({"bg_color": "#FFEB9C", "font_color": "#9C6500", "bold": True, "border": 1})
            })

        # 대상계약 서식
        ws_target = writer.sheets["대상계약"]
        ws_target.freeze_panes(1, 0)
        ws_target.autofilter(0, 0, max(len(target), 1), max(len(target.columns) - 1, 0))
        for col_idx, _ in enumerate(target.columns):
            ws_target.write(0, col_idx, target.columns[col_idx], fmt_header)
        autosize_columns(ws_target, target)
        for col_name in ["보험료", "가입금액", "월환산보험료", "연환산보험료", "CMP", "보험료_숫자", "인정보험료"]:
            if col_name in target.columns:
                idx = target.columns.get_loc(col_name)
                ws_target.set_column(idx, idx, 14, fmt_num)

        # 시책조건 서식
        ws_tier = writer.sheets["시책조건"]
        ws_tier.merge_range("A1:E1", "통합건강1 금시상 구간표", fmt_title)
        ws_tier.write("A2", "적용대상", fmt_header)
        ws_tier.write("B2", "새로담는건강보험 / 새로담는간편건강보험 / 새로담는건강보험플러스", fmt_body)
        ws_tier.write("A3", "기준", fmt_header)
        ws_tier.write("B3", "대상상품 합산 인정보험료 기준, 건강보험 건당 30만원 한도 반영", fmt_body)
        for col_idx, col_name in enumerate(tier_table.columns):
            ws_tier.write(3, col_idx, col_name, fmt_header)
        autosize_columns(ws_tier, tier_table)
        for col_name in ["달성구간", "13회차 금시상/현금"]:
            if col_name in tier_table.columns:
                idx = tier_table.columns.get_loc(col_name)
                ws_tier.set_column(idx, idx, 18, fmt_num)

    output.seek(0)
    return output.getvalue()


# -------------------------------
# Streamlit 화면
# -------------------------------

def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")
    gate_with_password()

    st.title(APP_TITLE)
    st.caption("엑셀 로우파일을 업로드하면 통합건강1 금시상 기준 콜리스트 엑셀을 자동으로 생성합니다.")

    with st.expander("현재 반영된 기준", expanded=True):
        st.markdown(
            """
            - 대상상품: **새로담는건강보험 / 새로담는간편건강보험 / 새로담는건강보험플러스**
            - 반영시책: **통합건강1 13회차 금시상 구간**
            - 인정기준: **계약상태=유지, 납입상태=정상, 건당 30만원 한도**
            - 구간: 5만원 / 10만원 / 20만원 / 30만원 / 50만원
            - 우선순위: S급 2만원 이하, A급 5만원 이하, B급 그 외
            """
        )

    uploaded = st.file_uploader("시책 로우 엑셀 파일 업로드", type=["xlsx", "xls"])
    if uploaded is None:
        st.info("엑셀 파일을 업로드하면 분석이 시작됩니다.")
        return

    try:
        xls = pd.ExcelFile(uploaded)
        sheet_name = st.selectbox("분석할 시트 선택", xls.sheet_names, index=0)
        uploaded.seek(0)
        df = read_excel_safely(uploaded, sheet_name)
    except Exception as exc:
        st.error(f"엑셀을 읽는 중 오류가 발생했습니다: {exc}")
        return

    missing = validate_columns(df)
    if missing:
        st.error(f"필수 컬럼이 없습니다: {', '.join(missing)}")
        st.stop()

    premium_candidates = [c for c in ["보험료", "월환산보험료", "CMP"] if c in df.columns]
    if not premium_candidates:
        numeric_like_cols = [c for c in df.columns if "보험" in c or "P" in c.upper() or "료" in c]
        premium_candidates = numeric_like_cols or list(df.columns)

    premium_col = st.selectbox("보험료 기준 컬럼", premium_candidates, index=0)
    col_a, col_b = st.columns(2)
    with col_a:
        filter_active = st.checkbox("계약상태=유지만 반영", value=True)
    with col_b:
        filter_normal_payment = st.checkbox("납입상태=정상만 반영", value=True)

    if st.button("콜리스트 생성하기", type="primary", use_container_width=True):
        try:
            call_list, target, metrics = analyze_promotion(df, premium_col=premium_col, filter_active=filter_active, filter_normal_payment=filter_normal_payment)
            excel_bytes = make_excel_file(call_list, target, metrics)
        except Exception as exc:
            st.error(f"분석 중 오류가 발생했습니다: {exc}")
            return

        st.success("콜리스트 생성이 완료되었습니다.")
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("대상 계약", f"{metrics['대상계약건수']:,}건")
        col2.metric("대상 설계사", f"{metrics['대상설계사수']:,}명")
        col3.metric("총 인정P", money(metrics["총인정보험료"]))
        col4.metric("S급", f"{metrics['S급']:,}명")
        col5.metric("A급", f"{metrics['A급']:,}명")

        st.subheader("콜리스트 미리보기")
        st.dataframe(call_list, use_container_width=True, hide_index=True)

        st.download_button(
            label="결과 엑셀 다운로드",
            data=excel_bytes,
            file_name="라이나_통합건강1_금시상_콜리스트.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        with st.expander("대상계약 미리보기"):
            st.dataframe(target, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
