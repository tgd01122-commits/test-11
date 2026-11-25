import hashlib

import pandas as pd
import streamlit as st

from industry_data import INDUSTRY_AVG, METRICS
from parsers import (ACCOUNTS, STATEMENTS, classify_sheet, detect_report_type, extract_from_excel,
                     extract_from_pdf, block_to_df, format_df, parse_excel, parse_pdf, sheet_label, sort_sheets)

st.set_page_config(page_title="사업보고서 분석기", layout="wide")


@st.cache_data
def cached_pdf(b):
    return parse_pdf(b)


@st.cache_data
def cached_excel(b):
    return parse_excel(b)


def ratio(num, den, as_pct=True):
    if num is None or den is None or den <= 0:
        return None
    r = num / den
    return r * 100 if as_pct else r


def compute_ratios(v):
    return {
        "부채비율": ratio(v.get("부채총계"), v.get("자본총계")),
        "유동비율": ratio(v.get("유동자산"), v.get("유동부채")),
        "자기자본비율": ratio(v.get("자본총계"), v.get("자산총계")),
        "영업이익률": ratio(v.get("영업이익"), v.get("매출액")),
        "순이익률": ratio(v.get("당기순이익"), v.get("매출액")),
        "ROE": ratio(v.get("당기순이익"), v.get("자본총계")),
        "ROA": ratio(v.get("당기순이익"), v.get("자산총계")),
        "총자산회전율": ratio(v.get("매출액"), v.get("자산총계"), as_pct=False),
    }


def judge(name, value, avg):
    gap = value - avg
    if avg and abs(gap / avg) <= 0.10:
        return "Neutral"
    if METRICS[name]["higher_better"]:
        return "Positive" if gap > 0 else "Negative"
    return "Positive" if gap < 0 else "Negative"


def fmt(value, name):
    if METRICS[name]["unit"] == "%":
        return f"{value:,.1f}%"
    return f"{value:,.2f}배"


def show_ratios(vals, kind):
    ratios = compute_ratios(vals)

    industry = st.selectbox("산업 선택", list(INDUSTRY_AVG.keys()), key=f"industry_{kind}")
    avg = INDUSTRY_AVG[industry]
    st.caption("산업 평균은 예시 값입니다.")

    verdicts = {}
    for name in METRICS:
        if ratios[name] is not None:
            verdicts[name] = judge(name, ratios[name], avg[name])

    counts = list(verdicts.values())
    c1, c2, c3 = st.columns(3)
    c1.metric("Positive", f"{counts.count('Positive')}개")
    c2.metric("Negative", f"{counts.count('Negative')}개")
    c3.metric("Neutral", f"{counts.count('Neutral')}개")

    st.divider()
    h1, h2 = st.columns(2)
    h1.subheader("기업")
    h2.subheader(f"{industry} 평균")

    for name, m in METRICS.items():
        with st.container(border=True):
            st.markdown(f"**{name}**")
            left, right = st.columns(2)
            right.metric("산업 평균", fmt(avg[name], name))
            if ratios[name] is None:
                left.metric("기업", "계산 불가")
                st.info("계산에 필요한 값이 없거나 0 이하입니다.")
                continue
            left.metric("기업", fmt(ratios[name], name))
            verdict = verdicts[name]
            if verdict == "Positive":
                st.success(f"Positive: {m['good']}")
            elif verdict == "Negative":
                st.error(f"Negative: {m['bad']}")
            else:
                st.warning("Neutral: 산업 평균과 비슷한 수준이에요.")

    pct_names = [n for n in METRICS if METRICS[n]["unit"] == "%" and ratios[n] is not None]
    if pct_names:
        chart = pd.DataFrame(
            {"기업": [ratios[n] for n in pct_names], f"{industry} 평균": [avg[n] for n in pct_names]},
            index=pct_names)
        st.bar_chart(chart)


def show_excel_sheets(sheets, file_id, kind):
    names = sort_sheets(sheets)
    fin_names = [n for n in names if classify_sheet(n, sheets[n])]
    if not fin_names:
        fin_names = names[:10]
    labels = []
    for n in fin_names:
        label = sheet_label(n, sheets[n])
        if label in labels:
            label = f"{label} {labels.count(label) + 1}"
        labels.append(label)
    for tab, n in zip(st.tabs(labels), fin_names):
        with tab:
            df = sheets[n].dropna(how="all").dropna(axis=1, how="all")
            if len(df) and str(df.iloc[0, 0]).strip() == "Index":
                df = df.iloc[1:]
            st.dataframe(format_df(df), hide_index=True)


def analyze(kind):
    quarterly = kind == "분기보고서"
    up = st.file_uploader(f"{kind} 업로드 (PDF, XLSX, XLS)", type=["pdf", "xlsx", "xls"], key=f"up_{kind}")
    scope = st.radio("재무제표 기준", ["연결", "별도"], horizontal=True, key=f"scope_{kind}")
    if up is None:
        st.info("파일을 올려 주세요.")
        return

    data = up.getvalue()
    file_id = hashlib.md5(data).hexdigest()[:8]
    is_pdf = up.name.lower().endswith(".pdf")

    try:
        if is_pdf:
            parsed = cached_pdf(data)
            extracted, used = extract_from_pdf(parsed, scope, quarterly)
            if detect_report_type(parsed) != ("분기" if quarterly else "사업"):
                other = "사업보고서" if quarterly else "분기보고서"
                st.warning(f"이 파일은 {other}로 보여요. {other} 탭에서 올리면 더 정확해요.")
        else:
            sheets = cached_excel(data)
            extracted, used = extract_from_excel(sheets, scope, quarterly)
    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했어요: {e}")
        return

    if used is None:
        st.warning("재무제표 제목을 찾지 못해 문서 전체에서 값을 찾았어요. 값이 틀릴 수 있으니 꼭 확인하세요.")
    elif used != scope:
        st.info(f"{scope} 재무제표를 찾지 못해 {used} 기준으로 읽었어요.")
    if quarterly:
        st.caption("분기보고서는 누적 값 기준으로 계산해요.")

    tab1, tab2 = st.tabs(["재무제표", "재무비율 비교"])

    with tab1:
        if is_pdf:
            blocks = parsed["blocks"].get(used, {}) if used else {}
            names = [n for n in STATEMENTS if n in blocks]
            if not names:
                st.warning("재무제표를 찾지 못했어요.")
            else:
                for tab, n in zip(st.tabs(names), names):
                    with tab:
                        st.dataframe(block_to_df(blocks[n]["text"]), hide_index=True, height=600)
                        with st.expander("원문 보기"):
                            st.text(blocks[n]["text"])
        else:
            show_excel_sheets(sheets, file_id, kind)

    with tab2:
        with st.expander("추출된 계정값 확인 및 수정"):
            vals = {}
            cols = st.columns(4)
            for i, acct in enumerate(ACCOUNTS):
                vals[acct] = cols[i % 4].number_input(
                    acct, value=float(extracted.get(acct) or 0.0), format="%.0f",
                    key=f"{kind}_{file_id}_{used}_{acct}")
        vals = {k: (v if v != 0 else None) for k, v in vals.items()}
        show_ratios(vals, kind)


with st.sidebar:
    st.header("DART")
    st.link_button("DART 전자공시", "https://dart.fss.or.kr")
    st.link_button("DART 상세검색", "https://dart.fss.or.kr/dsab007/main.do")

st.title("사업보고서 분석기")

annual_tab, quarter_tab = st.tabs(["사업보고서", "분기보고서"])
with annual_tab:
    analyze("사업보고서")
with quarter_tab:
    analyze("분기보고서")
