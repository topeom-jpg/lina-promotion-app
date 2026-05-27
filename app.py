from __future__ import annotations

from io import BytesIO
from typing import Iterable, Optional

import pandas as pd
import streamlit as st


APP_TITLE = "라이나 금시상 콜리스트 자동생성_엄성훈v2/26/05/27"

TARGET_PRODUCT_KEYWORDS = [
    "새로담는건강보험",
    "새로담는간편건강보험",
    "새로담는건강보험플러스",
]

PROMOTION_CONFIGS = {
    "1형": {
        "caption": "1형 통합건강1 금시상",
        "monthly_rate": "익월 150%",
        "max_rate": "13회차 최대 1,200% + @ / 금시상 최대 1,000%",
        "tiers": [
            {"tier": 50_000, "tier_name": "5만원 이상", "prize": 500_000, "prize_name": "50만원"},
            {"tier": 100_000, "tier_name": "10만원 이상", "prize": 1_000_000, "prize_name": "100만원"},
            {"tier": 200_000, "tier_name": "20만원 이상", "prize": 2_000_000, "prize_name": "200만원"},
            {"tier": 300_000, "tier_name": "30만원 이상", "prize": 3_000_000, "prize_name": "300만원"},
            {"tier": 500_000, "tier_name": "50만원 이상", "prize": 5_000_000, "prize_name": "500만원"},
        ],
    },
    "2형": {
        "caption": "2형 통합건강1 금시상",
        "monthly_rate": "익월 200%",
        "max_rate": "13회차 최대 2,400% + @ / 금시상 최대 2,000%",
        "tiers": [
            {"tier": 50_000, "tier_name": "5만원 이상", "prize": 600_000, "prize_name": "60만원"},
            {"tier": 100_000, "tier_name": "10만원 이상", "prize": 1_400_000, "prize_name": "140만원"},
            {"tier": 200_000, "tier_name": "20만원 이상", "prize": 3_200_000, "prize_name": "320만원"},
            {"tier": 300_000, "tier_name": "30만원 이상", "prize": 5_400_000, "prize_name": "540만원"},
            {"tier": 500_000, "tier_name": "50만원 이상", "prize": 10_000_000, "prize_name": "1,000만원"},
        ],
    },
}

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
    df.columns = [normalize_colname(c) for c in df.columns]
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


def current_tier_info(total_p: float, tiers: list[dict]) -> dict:
    current = {"tier": 0, "tier_name": "미달성", "prize": 0, "prize_name": "0원"}
    next_item: Optional[dict] = tiers[0]

    for item in tiers:
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


def make_subscription_note(row: pd.Series) -> str:
    try:
        subscription_count = int(row.get("청약계약건수", 0) or 0)
        subscription_p = int(row.get("청약인정보험료", 0) or 0)
    except Exception:
        subscription_count = 0
        subscription_p = 0

    if subscription_count <= 0 or subscription_p <= 0:
        return ""

    return (
        f"청약상태 계약 {subscription_count:,}건 / 인정P {subscription_p:,}원이 포함되어 있습니다. "
        "철회·거절 시 현재구간, 부족P, 금시상 대상 여부가 변동될 수 있으니 반드시 확인하세요."
    )


def make_call_message(row: pd.Series) -> str:
    promotion_type = row.get("시책유형", "")
    current_p = int(row["현재 통합건강1 P"])
    priority = row["콜우선순위"]
    subscription_note = str(row.get("청약확인멘트", "") or "")
    note_suffix = f" ※ {subscription_note}" if subscription_note else ""

    if priority == "완료":
        return f"[{promotion_type}] 현재 통합건강1 인정P {current_p:,}원으로 50만원 최고구간 달성입니다. 유지/철회 방어 체크가 우선입니다.{note_suffix}"

    shortage = int(row["부족P"])
    next_tier_name = row["다음구간명"]
    next_prize = row["다음금시상명"]
    increase = int(row["추가상승액"])

    if current_p == 0:
        return f"[{promotion_type}] 통합건강1 가동이 없습니다. {shortage:,}원 설계 시 {next_tier_name} 구간 진입, 13회차 금시상 {next_prize} 대상입니다.{note_suffix}"

    return (
        f"[{promotion_type}] 현재 통합건강1 인정P {current_p:,}원입니다. "
        f"{shortage:,}원만 추가하면 {next_tier_name} 구간 진입, "
        f"13회차 금시상 {next_prize} 대상입니다. 현재 대비 추가상승액은 {increase:,}원입니다."
        f"{note_suffix}"
    )

def get_agencies(df: pd.DataFrame, agency_col: str = "대리점명") -> list[str]:
    if agency_col not in df.columns:
        return []
    values = (
        df[agency_col]
        .fillna("")
        .astype(str)
        .str.strip()
    )
    agencies = sorted([v for v in values.unique().tolist() if v])
    return agencies


# -------------------------------
# 핵심 분석 로직
# -------------------------------

def analyze_promotion(
    df: pd.DataFrame,
    premium_col: str,
    promotion_type: str,
    filter_valid_status: bool = True,
    filter_normal_payment: bool = True,
) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    work = df.copy()
    work.columns = [normalize_colname(c) for c in work.columns]

    for col in ["대리점명", "지점명", "계약상태", "납입상태"]:
        if col not in work.columns:
            work[col] = ""

    if premium_col not in work.columns:
        raise ValueError(f"보험료 기준 컬럼 '{premium_col}'을 찾을 수 없습니다.")

    tiers = PROMOTION_CONFIGS[promotion_type]["tiers"]

    work["시책유형"] = promotion_type
    work["상품명"] = work["상품명"].fillna("").astype(str)
    work["보험료_숫자"] = to_number(work[premium_col])
    work["대상상품여부"] = contains_any(work["상품명"], TARGET_PRODUCT_KEYWORDS)

    if "계약상태" in work.columns:
        work["계약상태"] = work["계약상태"].fillna("").astype(str).str.strip()
    if "납입상태" in work.columns:
        work["납입상태"] = work["납입상태"].fillna("").astype(str).str.strip()

    status_mask = True
    if filter_valid_status and "계약상태" in work.columns:
        status_mask = work["계약상태"].isin(["유지", "청약"])
    payment_mask = True
    if filter_normal_payment and "납입상태" in work.columns:
        payment_mask = work["납입상태"].eq("정상")

    target = work[work["대상상품여부"] & status_mask & payment_mask].copy()
    target["인정보험료"] = target["보험료_숫자"].clip(upper=300_000)
    target["청약여부"] = target["계약상태"].eq("청약")
    target["청약보험료"] = target["보험료_숫자"].where(target["청약여부"], 0)
    target["청약인정보험료"] = target["인정보험료"].where(target["청약여부"], 0)
    target["청약확인대상"] = target["청약여부"].map(lambda x: "확인필요" if x else "")

    grouping_cols = ["시책유형", "대리점명", "지점명", "설계사"]
    summary = (
        target.groupby(grouping_cols, dropna=False)
        .agg(
            대상계약건수=("설계사", "size"),
            청약계약건수=("청약여부", "sum"),
            원보험료합계=("보험료_숫자", "sum"),
            청약보험료합계=("청약보험료", "sum"),
            **{"현재 통합건강1 P": ("인정보험료", "sum")},
            청약인정보험료=("청약인정보험료", "sum"),
        )
        .reset_index()
    )

    if summary.empty:
        summary = pd.DataFrame(columns=grouping_cols + ["대상계약건수", "청약계약건수", "원보험료합계", "청약보험료합계", "현재 통합건강1 P", "청약인정보험료"])

    if not summary.empty:
        tier_df = summary["현재 통합건강1 P"].apply(lambda x: current_tier_info(x, tiers)).apply(pd.Series)
        call_list = pd.concat([summary, tier_df], axis=1)
        call_list["청약확인멘트"] = call_list.apply(make_subscription_note, axis=1)
        call_list["콜멘트"] = call_list.apply(make_call_message, axis=1)
    else:
        call_list = summary.copy()
        for col in [
            "현재구간", "현재구간명", "다음구간", "다음구간명", "부족P",
            "현재금시상", "현재금시상명", "다음금시상", "다음금시상명", "추가상승액",
            "콜우선순위", "청약확인멘트", "콜멘트",
        ]:
            call_list[col] = []

    priority_order = {"S급": 1, "A급": 2, "B급": 3, "완료": 4}
    call_list["정렬키"] = call_list["콜우선순위"].map(priority_order).fillna(9)
    call_list = call_list.sort_values(
        by=["정렬키", "부족P", "추가상승액", "현재 통합건강1 P"],
        ascending=[True, True, False, False],
    ).drop(columns=["정렬키"]).reset_index(drop=True)

    target_display_cols = [c for c in [
        "시책유형", "NO", "대리점명", "지점명", "설계사", "계약번호", "계약일자", "계약상태", "상품명",
        "보험료", "계약자명", "피보험자명", "가입금액", "월환산보험료", "연환산보험료", "CMP",
        "납입기간", "납입주기", "납입상태", "청약확인대상", "보험료_숫자", "인정보험료", "청약인정보험료",
    ] if c in target.columns]
    target = target[target_display_cols].copy() if not target.empty else pd.DataFrame(columns=target_display_cols)

    metrics = {
        "시책유형": promotion_type,
        "대상계약건수": int(len(target)),
        "대상설계사수": int(call_list["설계사"].nunique()) if "설계사" in call_list.columns else 0,
        "총인정보험료": int(target["인정보험료"].sum()) if "인정보험료" in target.columns else 0,
        "청약계약건수": int(target["청약여부"].sum()) if "청약여부" in target.columns else 0,
        "청약원보험료": int(target["청약보험료"].sum()) if "청약보험료" in target.columns else 0,
        "청약인정보험료": int(target["청약인정보험료"].sum()) if "청약인정보험료" in target.columns else 0,
        "S급": int((call_list["콜우선순위"] == "S급").sum()) if "콜우선순위" in call_list.columns else 0,
        "A급": int((call_list["콜우선순위"] == "A급").sum()) if "콜우선순위" in call_list.columns else 0,
        "완료": int((call_list["콜우선순위"] == "완료").sum()) if "콜우선순위" in call_list.columns else 0,
    }
    return call_list, target, metrics


def split_by_agency_type(df: pd.DataFrame, agency_type_df: pd.DataFrame, agency_col: str = "대리점명") -> dict[str, pd.DataFrame]:
    if agency_col not in df.columns:
        selected_type = agency_type_df.loc[0, "시책유형"] if not agency_type_df.empty else "2형"
        return {selected_type: df.copy()}

    mapping = dict(zip(agency_type_df[agency_col].astype(str).str.strip(), agency_type_df["시책유형"].astype(str).str.strip()))
    work = df.copy()
    work["_대리점정리"] = work[agency_col].fillna("").astype(str).str.strip()
    work["_시책유형"] = work["_대리점정리"].map(mapping).fillna("2형")
    return {
        "1형": work[work["_시책유형"].eq("1형")].drop(columns=["_대리점정리", "_시책유형"]).copy(),
        "2형": work[work["_시책유형"].eq("2형")].drop(columns=["_대리점정리", "_시책유형"]).copy(),
    }


# -------------------------------
# 엑셀 생성 로직
# -------------------------------

def autosize_columns(worksheet, dataframe: pd.DataFrame, start_col: int = 0, max_width: int = 42) -> None:
    for idx, col in enumerate(dataframe.columns):
        if len(dataframe) == 0:
            width = len(str(col)) + 2
        else:
            values = dataframe[col].astype(str).replace("nan", "")
            width = max(len(str(col)), int(values.map(len).quantile(0.95)) if len(values) else 0) + 2
        worksheet.set_column(start_col + idx, start_col + idx, min(max(width, 10), max_width))


def tier_table_for_excel() -> pd.DataFrame:
    rows = []
    for ptype, config in PROMOTION_CONFIGS.items():
        for item in config["tiers"]:
            rows.append({
                "시책유형": ptype,
                "구분": config["caption"],
                "달성구간": item["tier"],
                "구간명": item["tier_name"],
                "13회차 금시상/현금": item["prize"],
                "지급액표시": item["prize_name"],
                "익월": config["monthly_rate"],
                "포스터표기": config["max_rate"],
            })
    return pd.DataFrame(rows)


def safe_sheet_name(name: str) -> str:
    return name[:31]


def call_list_export_view(call_list: pd.DataFrame) -> pd.DataFrame:
    """사용자에게 보여줄 콜리스트는 청약 관련 상세 숫자 컬럼을 빼고 청약확인멘트만 남깁니다."""
    call_cols = [
        "시책유형", "대리점명", "지점명", "설계사", "대상계약건수",
        "원보험료합계", "현재 통합건강1 P", "현재구간명",
        "다음구간명", "부족P", "현재금시상", "다음금시상", "추가상승액", "콜우선순위", "청약확인멘트", "콜멘트",
    ]
    return call_list[[c for c in call_cols if c in call_list.columns]].copy() if not call_list.empty else pd.DataFrame(columns=call_cols)


def write_formatted_table(writer, sheet_name: str, df: pd.DataFrame, workbook, fmt_header, fmt_num, fmt_note) -> None:
    df.to_excel(writer, sheet_name=safe_sheet_name(sheet_name), index=False, startrow=0)
    ws = writer.sheets[safe_sheet_name(sheet_name)]
    ws.freeze_panes(1, 0)
    if len(df.columns) > 0:
        ws.autofilter(0, 0, max(len(df), 1), max(len(df.columns) - 1, 0))
        for col_idx, col_name in enumerate(df.columns):
            ws.write(0, col_idx, col_name, fmt_header)
        autosize_columns(ws, df)
        number_cols = [
            "대상계약건수", "청약계약건수", "원보험료합계", "청약보험료합계", "현재 통합건강1 P", "청약인정보험료",
            "현재구간", "다음구간", "부족P", "현재금시상", "다음금시상", "추가상승액",
            "보험료", "가입금액", "월환산보험료", "연환산보험료", "CMP", "보험료_숫자", "인정보험료",
            "달성구간", "13회차 금시상/현금",
        ]
        for col_name in number_cols:
            if col_name in df.columns:
                idx = df.columns.get_loc(col_name)
                ws.set_column(idx, idx, 16, fmt_num)
        if "청약확인멘트" in df.columns:
            idx = df.columns.get_loc("청약확인멘트")
            ws.set_column(idx, idx, 70, fmt_note)
            if len(df) > 0:
                ws.conditional_format(1, idx, len(df), idx, {
                    "type": "text", "criteria": "containing", "value": "청약상태",
                    "format": workbook.add_format({"bg_color": "#FFF200", "font_color": "#7F6000", "bold": True, "border": 1, "text_wrap": True}),
                })
        if "콜멘트" in df.columns:
            idx = df.columns.get_loc("콜멘트")
            ws.set_column(idx, idx, 75, fmt_note)
        if "콜우선순위" in df.columns and len(df) > 0:
            priority_col = df.columns.get_loc("콜우선순위")
            ws.conditional_format(1, priority_col, len(df), priority_col, {
                "type": "text", "criteria": "containing", "value": "S급",
                "format": workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006", "bold": True, "border": 1}),
            })
            ws.conditional_format(1, priority_col, len(df), priority_col, {
                "type": "text", "criteria": "containing", "value": "A급",
                "format": workbook.add_format({"bg_color": "#FFEB9C", "font_color": "#9C6500", "bold": True, "border": 1}),
            })


def make_excel_file_multi(results: dict[str, dict], agency_type_df: pd.DataFrame, premium_col: str) -> bytes:
    output = BytesIO()

    call_cols = [
        "시책유형", "대리점명", "지점명", "설계사", "대상계약건수",
        "원보험료합계", "현재 통합건강1 P", "현재구간명",
        "다음구간명", "부족P", "현재금시상", "다음금시상", "추가상승액", "콜우선순위", "청약확인멘트", "콜멘트",
    ]

    dashboard_rows = []
    for ptype in ["1형", "2형"]:
        metrics = results.get(ptype, {}).get("metrics", {})
        dashboard_rows.append({
            "시책유형": ptype,
            "대상계약건수": metrics.get("대상계약건수", 0),
            "대상설계사수": metrics.get("대상설계사수", 0),
            "총인정보험료": metrics.get("총인정보험료", 0),
            "청약계약건수": metrics.get("청약계약건수", 0),
            "청약인정보험료": metrics.get("청약인정보험료", 0),
            "S급": metrics.get("S급", 0),
            "A급": metrics.get("A급", 0),
            "완료": metrics.get("완료", 0),
        })
    dashboard_df = pd.DataFrame(dashboard_rows)
    tier_table = tier_table_for_excel()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
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
        dash.merge_range("A1:H1", "2026년 5월 라이나생명 통합건강1 금시상 콜작업용", fmt_title)
        dash.merge_range("A2:H2", "로우파일 업로드 후 대리점별 1형/2형을 선택하여 시트별 콜리스트 생성", fmt_subtitle)
        dash.write("A4", "분석 기준", fmt_header)
        dash.write("B4", "내용", fmt_header)
        base_rules = [
            ("대상 상품군", "새로담는건강보험 / 새로담는간편건강보험 / 새로담는건강보험플러스"),
            ("적용 시책", "통합건강1 13회차 금시상 구간만 반영"),
            ("보험료 기준 컬럼", premium_col),
            ("인정 기준", "계약상태=유지 또는 청약, 납입상태=정상, 건당 30만원 한도"),
            ("청약 포함 주의", "청약상태 계약이 포함된 설계사는 청약확인멘트에 표시됩니다. 철회·거절 시 구간 변동 여부를 확인하세요."),
            ("콜 등급", "S급: 부족P 2만원 이하 / A급: 5만원 이하 / B급: 그 외 / 완료: 50만원 이상"),
        ]
        for idx, (label, value) in enumerate(base_rules, start=5):
            dash.write(idx - 1, 0, label, fmt_body)
            dash.write(idx - 1, 1, value, fmt_body)

        dash.write("D4", "시책유형", fmt_header)
        dash.write("E4", "대상계약", fmt_header)
        dash.write("F4", "설계사", fmt_header)
        dash.write("G4", "총 인정P", fmt_header)
        dash.write("H4", "청약건수", fmt_header)
        dash.write("I4", "청약 인정P", fmt_header)
        dash.write("J4", "S/A급", fmt_header)
        for row_idx, row in enumerate(dashboard_df.itertuples(index=False), start=5):
            dash.write(row_idx - 1, 3, row.시책유형, fmt_metric_label)
            dash.write(row_idx - 1, 4, row.대상계약건수, fmt_metric_value)
            dash.write(row_idx - 1, 5, row.대상설계사수, fmt_metric_value)
            dash.write(row_idx - 1, 6, row.총인정보험료, fmt_metric_value)
            dash.write(row_idx - 1, 7, row.청약계약건수, fmt_metric_value)
            dash.write(row_idx - 1, 8, row.청약인정보험료, fmt_metric_value)
            dash.write(row_idx - 1, 9, f"S {row.S급:,} / A {row.A급:,}", fmt_body)

        all_calls = []
        for ptype in ["1형", "2형"]:
            call_list = results.get(ptype, {}).get("call_list", pd.DataFrame())
            if not call_list.empty:
                all_calls.append(call_list)
        if all_calls:
            combined = pd.concat(all_calls, ignore_index=True)
            top = combined[combined["콜우선순위"].isin(["S급", "A급"])].head(15)
            if top.empty:
                top = combined.head(15)
        else:
            top = pd.DataFrame(columns=["시책유형", "콜우선순위", "대리점명", "지점명", "설계사", "현재 통합건강1 P", "부족P", "다음구간명", "다음금시상명", "청약확인멘트"])
        top_cols = ["시책유형", "콜우선순위", "대리점명", "지점명", "설계사", "현재 통합건강1 P", "부족P", "다음구간명", "다음금시상명", "청약확인멘트"]
        top = top[[c for c in top_cols if c in top.columns]].copy()
        dash.write("A11", "우선 콜 TOP", fmt_header)
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
        dash.set_column("D:J", 16)
        dash.set_row(0, 28)

        # 대리점 분류표
        agency_export = agency_type_df.copy()
        write_formatted_table(writer, "대리점_분류표", agency_export, workbook, fmt_header, fmt_num, fmt_note)

        # 1형/2형 결과 시트
        for ptype in ["1형", "2형"]:
            call_list = results.get(ptype, {}).get("call_list", pd.DataFrame())
            target = results.get(ptype, {}).get("target", pd.DataFrame())
            call_export = call_list_export_view(call_list)
            write_formatted_table(writer, f"{ptype}_콜리스트", call_export, workbook, fmt_header, fmt_num, fmt_note)
            write_formatted_table(writer, f"{ptype}_대상계약", target, workbook, fmt_header, fmt_num, fmt_note)

        # 시책조건
        write_formatted_table(writer, "시책조건", tier_table, workbook, fmt_header, fmt_num, fmt_note)
        ws_tier = writer.sheets["시책조건"]
        ws_tier.write("J1", "참고", fmt_header)
        ws_tier.write("J2", "1형/2형 모두 통합건강1 금시상만 반영. 전략특약/연속가동/익월시책은 현재 버전 제외.", fmt_body)
        ws_tier.set_column("J:J", 80)

    output.seek(0)
    return output.getvalue()


# -------------------------------
# Streamlit 화면
# -------------------------------

def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide")
    gate_with_password()

    st.title(APP_TITLE)
    st.caption("엑셀 로우파일을 업로드하면 대리점별 1형/2형을 선택하고 통합건강1 금시상 콜리스트를 자동 생성합니다.")

    with st.expander("현재 반영된 기준", expanded=True):
        st.markdown(
            """
            - 대상상품: **새로담는건강보험 / 새로담는간편건강보험 / 새로담는건강보험플러스**
            - 반영시책: **통합건강1 13회차 금시상 구간만 반영**
            - 1형 금시상: **5만 50만원 / 10만 100만원 / 20만 200만원 / 30만 300만원 / 50만 500만원**
            - 2형 금시상: **5만 60만원 / 10만 140만원 / 20만 320만원 / 30만 540만원 / 50만 1,000만원**
            - 인정기준: **계약상태=유지 또는 청약, 납입상태=정상, 건당 30만원 한도**
            - 청약주의: **청약상태 계약은 합산하되, 결과 엑셀 콜리스트에는 청약확인멘트만 표시**
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
        filter_valid_status = st.checkbox("계약상태=유지/청약만 반영", value=True)
    with col_b:
        filter_normal_payment = st.checkbox("납입상태=정상만 반영", value=True)

    st.divider()
    st.subheader("대리점별 시책유형 선택")

    agencies = get_agencies(df, "대리점명")
    if agencies:
        st.caption("기본값은 전체 2형입니다. 1형 대리점만 체크해주세요. 체크하지 않은 대리점은 자동으로 2형 처리됩니다.")

        default_agency_check_df = pd.DataFrame({
            "1형 선택": [False] * len(agencies),
            "대리점명": agencies,
        })

        edited_agency_df = st.data_editor(
            default_agency_check_df,
            hide_index=True,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "1형 선택": st.column_config.CheckboxColumn("1형 선택", help="1형 대리점이면 체크", default=False),
                "대리점명": st.column_config.TextColumn("대리점명", disabled=True),
            },
            key="agency_type_checkbox_editor",
        )

        edited_agency_df["1형 선택"] = edited_agency_df["1형 선택"].fillna(False).astype(bool)
        agency_type_df = edited_agency_df[["대리점명", "1형 선택"]].copy()
        agency_type_df["시책유형"] = agency_type_df["1형 선택"].map(lambda checked: "1형" if checked else "2형")
        agency_type_df = agency_type_df[["대리점명", "시책유형"]]

        type1_agencies = agency_type_df.loc[agency_type_df["시책유형"].eq("1형"), "대리점명"].tolist()
        type2_agencies = agency_type_df.loc[agency_type_df["시책유형"].eq("2형"), "대리점명"].tolist()

        col_type1, col_type2 = st.columns(2)
        with col_type1:
            st.markdown("**1형 선택 대리점**")
            if type1_agencies:
                st.write(", ".join(type1_agencies))
            else:
                st.info("아직 1형으로 선택한 대리점이 없습니다.")
        with col_type2:
            st.markdown("**2형 자동 처리 대리점**")
            if type2_agencies:
                st.write(", ".join(type2_agencies))
            else:
                st.info("모든 대리점이 1형으로 선택되었습니다.")
    else:
        selected_type = st.radio("대리점명 컬럼이 없어 전체 파일에 적용할 시책유형을 선택하세요.", ["1형", "2형"], horizontal=True, index=1)
        agency_type_df = pd.DataFrame({"대리점명": ["전체"], "시책유형": [selected_type]})

    output_mode = st.radio(
        "결과 생성 방식",
        ["1형+2형 전체 생성", "1형만 생성", "2형만 생성"],
        horizontal=True,
        index=0,
    )

    if st.button("콜리스트 생성하기", type="primary", use_container_width=True):
        try:
            split_data = split_by_agency_type(df, agency_type_df, "대리점명")
            run_types = ["1형", "2형"]
            if output_mode == "1형만 생성":
                run_types = ["1형"]
            elif output_mode == "2형만 생성":
                run_types = ["2형"]

            results = {}
            for ptype in run_types:
                part_df = split_data.get(ptype, pd.DataFrame(columns=df.columns))
                call_list, target, metrics = analyze_promotion(
                    part_df,
                    premium_col=premium_col,
                    promotion_type=ptype,
                    filter_valid_status=filter_valid_status,
                    filter_normal_payment=filter_normal_payment,
                )
                results[ptype] = {"call_list": call_list, "target": target, "metrics": metrics}

            # 생성하지 않은 유형도 빈 시트/대시보드가 안정적으로 생기도록 빈 결과를 넣습니다.
            for ptype in ["1형", "2형"]:
                if ptype not in results:
                    empty_call, empty_target, empty_metrics = analyze_promotion(
                        pd.DataFrame(columns=df.columns),
                        premium_col=premium_col,
                        promotion_type=ptype,
                        filter_valid_status=filter_valid_status,
                        filter_normal_payment=filter_normal_payment,
                    )
                    results[ptype] = {"call_list": empty_call, "target": empty_target, "metrics": empty_metrics}

            excel_bytes = make_excel_file_multi(results, agency_type_df, premium_col)
        except Exception as exc:
            st.error(f"분석 중 오류가 발생했습니다: {exc}")
            return

        st.success("콜리스트 생성이 완료되었습니다.")
        m1 = results["1형"]["metrics"]
        m2 = results["2형"]["metrics"]
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        col1.metric("1형 대상계약", f"{m1['대상계약건수']:,}건")
        col2.metric("1형 S/A급", f"S {m1['S급']:,} / A {m1['A급']:,}")
        col3.metric("1형 총 인정P", money(m1["총인정보험료"]))
        col4.metric("2형 대상계약", f"{m2['대상계약건수']:,}건")
        col5.metric("2형 S/A급", f"S {m2['S급']:,} / A {m2['A급']:,}")
        col6.metric("2형 총 인정P", money(m2["총인정보험료"]))

        total_subscription_count = int(m1.get("청약계약건수", 0)) + int(m2.get("청약계약건수", 0))
        total_subscription_p = int(m1.get("청약인정보험료", 0)) + int(m2.get("청약인정보험료", 0))
        if total_subscription_count > 0:
            st.warning(
                f"청약상태 계약 {total_subscription_count:,}건 / 인정P {total_subscription_p:,}원이 합산되었습니다. "
                "철회·거절 시 구간과 금시상 대상 여부가 달라질 수 있으니 결과 엑셀의 청약확인멘트를 확인하세요."
            )

        tab1, tab2, tab3 = st.tabs(["1형 콜리스트", "2형 콜리스트", "대리점 분류표"])
        with tab1:
            st.dataframe(call_list_export_view(results["1형"]["call_list"]), use_container_width=True, hide_index=True)
        with tab2:
            st.dataframe(call_list_export_view(results["2형"]["call_list"]), use_container_width=True, hide_index=True)
        with tab3:
            st.dataframe(agency_type_df, use_container_width=True, hide_index=True)

        st.download_button(
            label="결과 엑셀 다운로드",
            data=excel_bytes,
            file_name="라이나_통합건강1_금시상_1형2형_콜리스트.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        with st.expander("대상계약 미리보기"):
            preview_target = pd.concat([results["1형"]["target"], results["2형"]["target"]], ignore_index=True)
            st.dataframe(preview_target, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
