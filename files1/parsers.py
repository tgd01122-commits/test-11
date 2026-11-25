import io
import re

import pandas as pd


STATEMENTS = {
    "재무상태표": dict(keys=["재무상태표"], sig=["자산총계"]),
    "손익계산서": dict(keys=["손익계산서"], sig=["영업이익", "당기순이익", "매출액"]),
    "현금흐름표": dict(keys=["현금흐름표"], sig=["영업활동"]),
    "자본변동표": dict(keys=["자본변동표"], sig=["기초"]),
}


ACCOUNTS = {
    "자산총계":   (["자산총계"], "재무상태표"),
    "부채총계":   (["부채총계"], "재무상태표"),
    "자본총계":   (["자본총계"], "재무상태표"),
    "유동자산":   (["유동자산"], "재무상태표"),
    "유동부채":   (["유동부채"], "재무상태표"),
    "매출액":     (["매출액", "영업수익", "수익(매출액)", "수익"], "손익계산서"),
    "영업이익":   (["영업이익(손실)", "영업이익", "영업손익", "영업손실"], "손익계산서"),
    "당기순이익": (["당기순이익(손실)", "당기순이익", "연결당기순이익", "분기순이익", "반기순이익", "당기순손익", "당기순손실", "분기순손실", "반기순손실"], "손익계산서"),
}

LOSS_ALIASES = {"영업손실", "당기순손실", "분기순손실", "반기순손실"}

_PREFIX = r"^\s*(?:[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVX\d]{1,4}[\.\)]\s*)?"
_NUM = re.compile(r"[\(△▲-]?\d[\d,\.]*\)?")


def to_num(x):
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return None if pd.isna(x) else float(x)
    s = str(x).strip().replace(",", "").replace(" ", "")
    if not s:
        return None
    neg = s.startswith("(") or s.startswith("△") or s.startswith("▲") or s.startswith("-")
    s = s.strip("()△▲-")
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def _label_regex(alias):
    body = r"\s*".join(re.escape(c) for c in alias)
    return re.compile(_PREFIX + body + r"(?![가-힣])")


def _apply_loss(alias, v):
    if alias in LOSS_ALIASES:
        return -abs(v)
    return v


def _pick_number(tokens, col=0):
    vals = []
    for t in tokens:
        digits = re.sub(r"\D", "", t)
        if "," in t or len(digits) >= 4:
            v = to_num(t)
            if v is not None:
                vals.append(v)
    if not vals:
        return None
    return vals[col] if len(vals) > col else vals[0]


def parse_excel(file_bytes):
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None)


def sheet_title(name, df=None):
    if df is not None:
        try:
            t = str(df.iloc[1, 0])
            if t.startswith("["):
                return t
        except Exception:
            pass
    return str(name)


def classify_sheet(name, df=None):
    t = sheet_title(name, df).replace(" ", "")
    for key, cfg in STATEMENTS.items():
        if any(k in t for k in cfg["keys"]):
            return key
    return None


def sheet_label(name, df=None):
    t = re.sub(r"^\[[^\]]*\]\s*", "", sheet_title(name, df))
    kind = classify_sheet(name, df)
    if not kind:
        return str(name)
    base = "포괄손익계산서" if kind == "손익계산서" and "포괄" in t else kind
    scope = "연결" if "연결" in t else "별도" if "별도" in t else ""
    return f"{base} ({scope})" if scope else base


def sort_sheets(sheets):
    order = list(STATEMENTS.keys())

    def k(n):
        c = classify_sheet(n, sheets[n])
        t = sheet_title(n, sheets[n])
        return (order.index(c) if c else 99, 0 if "연결" in t else 1)

    return sorted(sheets.keys(), key=k)


def format_df(df):
    def f(x):
        if isinstance(x, bool):
            return str(x)
        if isinstance(x, (int, float)):
            if pd.isna(x):
                return ""
            return f"{int(x):,}" if float(x).is_integer() else f"{x:,.2f}"
        if x is None or (not isinstance(x, str) and pd.isna(x)):
            return ""
        return str(x)

    out = df.astype(object).apply(lambda col: col.map(f))
    out.columns = [" " * (i + 1) for i in range(out.shape[1])]
    return out


def _sheet_cumulative(df):
    head = " ".join(str(x) for x in df.head(10).values.flatten() if isinstance(x, str))
    flat = head.replace(" ", "")
    return "3개월" in flat and "누적" in flat


def _value_from_sheet(df, aliases, col=0):
    for alias in aliases:
        reg = _label_regex(alias)
        for _, row in df.iterrows():
            cells = list(row.values)
            for i, c in enumerate(cells[:3]):
                if isinstance(c, str) and reg.search(c):
                    nums = [to_num(x) for x in cells[i + 1:]]
                    nums = [n for n in nums if n is not None]
                    big = [n for n in nums if abs(n) >= 1000]
                    pool = big if big else nums
                    if pool:
                        v = pool[col] if len(pool) > col else pool[0]
                        return _apply_loss(alias, v)
    return None


def extract_from_excel(sheets, scope="연결", quarterly=False):
    out = {}
    other = "별도" if scope == "연결" else "연결"
    names = sort_sheets(sheets)
    stmt_sheets = [n for n in names if classify_sheet(n, sheets[n])]
    used = scope if any(scope in sheet_title(n, sheets[n]) for n in stmt_sheets) else other
    names = sorted(names, key=lambda n: used not in sheet_title(n, sheets[n]))
    for acct, (aliases, stmt) in ACCOUNTS.items():
        same = [n for n in names if classify_sheet(n, sheets[n]) == stmt]
        targets = same + [n for n in names if n not in same]
        for n in targets:
            col = 1 if quarterly and n in same and stmt == "손익계산서" and _sheet_cumulative(sheets[n]) else 0
            v = _value_from_sheet(sheets[n], aliases, col)
            if v is not None:
                out[acct] = v
                break
    return _fill_derived(out), used


_TITLE = re.compile(
    r"^(?:[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVX]+[\.\)])?(?:\d+(?:-\d+)*[\.\)])?(?:\(\d+\))?"
    r"(?:요약)?(?:분기|반기)?(연결|별도)?(재무상태표|포괄손익계산서|손익계산서|현금흐름표|자본변동표)$"
)


def _is_valid_block(kind, text):
    flat = text.replace(" ", "")
    if kind == "재무상태표":
        return "자산총계" in flat
    if kind == "손익계산서":
        return any(k in flat for k in ("매출액", "영업이익", "영업손실", "영업수익"))
    if kind == "현금흐름표":
        return "영업활동" in flat
    return "기초" in flat or "기말" in flat


def _find_blocks(pages):
    lines = []
    for page_no, text in enumerate(pages):
        for line in text.splitlines():
            lines.append((page_no, line))
    titles = []
    for i, (_, line) in enumerate(lines):
        m = _TITLE.match(re.sub(r"\s+", "", line))
        if m:
            kind = "손익계산서" if "손익계산서" in m.group(2) else m.group(2)
            titles.append((i, m.group(1) or "별도", kind))
    blocks = {"연결": {}, "별도": {}}
    for n, (start, scope, kind) in enumerate(titles):
        if kind in blocks[scope]:
            continue
        end = len(lines)
        for next_start, next_scope, next_kind in titles[n + 1:]:
            if (next_scope, next_kind) != (scope, kind):
                end = next_start
                break
        chunk = lines[start:min(end, start + 400)]
        text = "\n".join(line for _, line in chunk)
        if _is_valid_block(kind, text):
            blocks[scope][kind] = dict(text=text, pages=sorted({p + 1 for p, _ in chunk}))
    return blocks


def parse_pdf(file_bytes):
    import pdfplumber

    pages = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for p in pdf.pages:
            pages.append(p.extract_text() or "")
    return dict(pages=pages, blocks=_find_blocks(pages))


def detect_report_type(parsed):
    head = "".join("".join(parsed["pages"][:3]).split())
    if any(k in head for k in ("분기보고서", "반기보고서", "분기검토", "반기검토", "분기재무제표", "반기재무제표")):
        return "분기"
    return "사업"


def _is_cumulative(text):
    flat = text.replace(" ", "")
    return "3개월" in flat and "누적" in flat


def _value_from_text(text, aliases, col=0):
    for alias in aliases:
        reg = _label_regex(alias)
        for line in text.splitlines():
            m = reg.search(line)
            if m:
                rest = re.sub(r"\(주[^)]*\)", "", line[m.end():])
                v = _pick_number(_NUM.findall(rest), col)
                if v is not None:
                    return _apply_loss(alias, v)
    return None


def select_blocks(parsed, scope):
    other = "별도" if scope == "연결" else "연결"
    for sc in (scope, other):
        blocks = parsed["blocks"][sc]
        if "재무상태표" in blocks and "손익계산서" in blocks:
            return sc, blocks
    return None, {}


def extract_from_pdf(parsed, scope="연결", quarterly=False):
    out = {}
    used, blocks = select_blocks(parsed, scope)
    if not blocks:
        whole = "\n".join(parsed["pages"])
        for acct, (aliases, stmt) in ACCOUNTS.items():
            v = _value_from_text(whole, aliases)
            if v is not None:
                out[acct] = v
        return _fill_derived(out), None
    header = "\n".join(blocks["손익계산서"]["text"].splitlines()[:12])
    income_col = 1 if quarterly and _is_cumulative(header) else 0
    for acct, (aliases, stmt) in ACCOUNTS.items():
        col = income_col if stmt == "손익계산서" else 0
        v = _value_from_text(blocks[stmt]["text"], aliases, col)
        if v is not None:
            out[acct] = v
    return _fill_derived(out), used


_VALUE = r"(?:[\(△▲-]?\d[\d,\.]*\)?|-)"
_ROW = re.compile(r"^(.*?)\s+(" + _VALUE + r"(?:\s+" + _VALUE + r")*)\s*$")


def block_to_df(text):
    rows = []
    for line in text.splitlines():
        line = line.strip()
        if not line or "dart.fss.or.kr" in line:
            continue
        m = _ROW.match(line)
        if m and m.group(1).strip():
            rows.append((m.group(1).strip(), m.group(2).split()))
        else:
            rows.append((line, []))
    n = max([len(v) for _, v in rows] + [1])
    if n == 4 and _is_cumulative(text):
        names = ["당기 3개월", "당기 누적", "전기 3개월", "전기 누적"]
    elif n == 2:
        names = ["당기", "전기"]
    else:
        names = [f"값{i + 1}" for i in range(n)]
    data = [[label] + vals + [""] * (n - len(vals)) for label, vals in rows]
    return pd.DataFrame(data, columns=["과목"] + names)


def _fill_derived(v):
    if "부채총계" not in v and "자산총계" in v and "자본총계" in v:
        v["부채총계"] = v["자산총계"] - v["자본총계"]
    return v
