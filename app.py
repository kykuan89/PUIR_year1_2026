import re
import pandas as pd
import streamlit as st
import plotly.express as px
from normalization import (
    infer_question_type,
    normalize_column_single,
    normalize_multiselect_cell,
    NORMALIZE_MAP,
    parse_likert_1to5_unknown6,
    add_college_column
)

FILE_PATH = "114學年度填答總表_含班級學號12102025_去識別化.xlsx"
SHEET = "1210-已填總表"

# ---------- Data loading ----------
@st.cache_data
def load_data(path: str = FILE_PATH, sheet: str = SHEET):
    df = pd.read_excel(path, sheet_name=sheet)
    df = df.dropna(axis=1, how="all")

    obj_cols = df.select_dtypes(include=["object", "string"]).columns
    df[obj_cols] = df[obj_cols].apply(lambda s: s.astype("string").str.strip())

    auto_cols = auto_detect_likert_1to5_cols(df)

    # 一次建立 Likert 衍生欄位
    likert_new_cols = {}
    for col in auto_cols:
        parsed = df[col].apply(parse_likert_1to5_unknown6)
        likert_new_cols[col + "__num"] = parsed.apply(lambda t: t[0])
        likert_new_cols[col + "__cat"] = parsed.apply(lambda t: t[1])

    if likert_new_cols:
        df = pd.concat([df, pd.DataFrame(likert_new_cols, index=df.index)], axis=1)

    exclude_cols = set(["已填人"])
    schema = {}
    for col in df.columns:
        if col in exclude_cols:
            continue

        rule = NORMALIZE_MAP.get(col)
        if rule and rule.get("type") == "multiselect":
            schema[col] = {"type": "multiselect"}
        else:
            schema[col] = {"type": infer_question_type(col, df[col])}

    # 一次處理 normalization
    replace_cols = {}
    norm_new_cols = {}

    for col, rule in NORMALIZE_MAP.items():
        if col not in df.columns:
            continue

        qtype = rule.get("type") or schema.get(col, {}).get("type", "single")

        if qtype == "multiselect":
            norm_new_cols[col + "_norm"] = df[col].apply(lambda x: normalize_multiselect_cell(x, rule))
        else:
            replace_cols[col] = normalize_column_single(df[col], rule)

    for col, series in replace_cols.items():
        df[col] = series

    if norm_new_cols:
        df = pd.concat([df, pd.DataFrame(norm_new_cols, index=df.index)], axis=1)

    df = df.copy()

    return df, schema

NUM_PREFIX = re.compile(r"^\s*(\d+)\s*(.*)$", re.I)

def auto_detect_likert_1to5_cols(df: pd.DataFrame, min_15_unique: int = 3) -> list[str]:
    cols = []
    obj_cols = df.select_dtypes(include=["object", "string"]).columns

    for col in obj_cols:
        s = df[col].dropna().astype(str).str.strip()
        if s.empty:
            continue

        codes = set()
        total = len(s)

        for v in s:
            m = NUM_PREFIX.match(v)
            if m:
                code = int(m.group(1))
                if 1 <= code <= 5:
                    codes.add(code)

        # ✅ 條件1：至少有幾種 Likert 分數
        cond1 = len(codes) >= min_15_unique

        # ✅ 條件2：大部分值是 1~5（避免誤判）
        likert_ratio = sum(
            1 for v in s
            if (NUM_PREFIX.match(v) and 1 <= int(NUM_PREFIX.match(v).group(1)) <= 5)
        ) / total

        cond2 = likert_ratio > 0.5

        if cond1 and cond2:
            cols.append(col)

    return cols

def is_multiselect(col_name: str) -> bool:
    return "可複選" in str(col_name)

def split_multiselect(series: pd.Series) -> pd.Series:
    # 支援 ; ； , 、 等分隔（依你資料狀況可再加）
    s = series.dropna().astype(str)
    s = s.str.replace(r"\s+", " ", regex=True).str.strip()
    parts = s.str.split(r"\s*[;；,、]\s*", regex=True)
    return parts.explode().str.strip()

def summarize(df: pd.DataFrame, schema: dict, q: str, group: str | None, as_percent: bool) -> pd.DataFrame:
    q_use = q + "__cat" if (q + "__cat") in df.columns else q
    d = df.copy()

    is_ms = (schema.get(q, {}).get("type") == "multiselect")

    if is_ms:
        coln = q + "_norm"
        if coln in d.columns:
            exploded = d[coln].explode()
            d = d.loc[exploded.index].copy()
            d[q_use] = exploded
        else:
            exploded = split_multiselect(d[q])
            d = d.loc[exploded.index].copy()
            d[q_use] = exploded

        d = d[d[q_use].notna() & (d[q_use].astype(str).str.strip() != "")]
    else:
        d = d[d[q_use].notna() & (d[q_use].astype(str).str.strip() != "")]

    if group and group != "(不分組)":
        d = d[d[group].notna() & (d[group].astype(str).str.strip() != "")]

        ct = pd.crosstab(d[group], d[q_use])
        out = ct.stack().reset_index()
        out.columns = [group, q_use, "count"]

        if as_percent:
            if is_ms:
                n_resp = d.groupby(group).apply(lambda g: g.index.nunique())
                out["percent"] = out["count"] / out[group].map(n_resp) * 100
            else:
                # 單選題維持「各組加總=100%」
                out["percent"] = out.groupby(group)["count"].transform(lambda x: x / x.sum() * 100)

        return out
    else:
        vc = d[q_use].value_counts(dropna=False).rename_axis(q_use).reset_index(name="count")

        if as_percent:
            if is_ms:
                # ✅ 不分組：分母用作答人數
                n_resp = d.index.nunique()
                vc["percent"] = vc["count"] / n_resp * 100
            else:
                vc["percent"] = vc["count"] / vc["count"].sum() * 100

        return vc



def get_plot_series(df, schema, q):
    # multiselect：用 _norm
    if schema.get(q, {}).get("type") == "multiselect":
        coln = q + "_norm"
        if coln in df.columns:
            return df[coln].explode()
        # 萬一沒產生 _norm，退回原欄位（但會很亂）
        return df[q].astype("string")
    # single：用原欄位
    return df[q].astype("string")


# ---------- UI ----------
st.set_page_config(page_title="115學年度大一新生學習適應性分析", layout="wide")
st.title("115學年度大一新生學習適應性分析2")

df, schema = load_data(FILE_PATH)

df = add_college_column(df)
all_questions = [
    c for c in df.columns
    if c not in ["已填人"]
    and not c.endswith("__num")
    and not c.endswith("__cat")
    and not c.endswith("_norm")
]

default_groups = ["(不分組)"]
candidate_groups = [
    c for c in [
        "1性別",
        "2身分別",
        "學院",         
        "班級",
        "5原畢業學校之類型",
        "6原畢業學校所在地區",
    ]
    if c in df.columns
]
group_options = default_groups + candidate_groups

with st.sidebar:
    st.header("設定")
    q = st.selectbox("選擇題目欄位", all_questions)
    group = st.selectbox("分組比較", group_options)
    as_percent = st.checkbox("顯示百分比 (%)", value=True)

result = summarize(df, schema, q=q, group=None if group == "(不分組)" else group, as_percent=as_percent)

q_base = q.replace("__cat", "").replace("__num", "")
st.subheader(f"題目：{q_base}")

# ---- 顯示 Likert 平均與標準差（只要有 1~5 就顯示）----
num_col = f"{q_base}__num"

if num_col in df.columns:
    raw = pd.to_numeric(df[num_col], errors="coerce")
    v = raw.dropna()

    total_n = len(raw)
    valid_n = len(v)
    unknown_n = total_n - valid_n

    if valid_n > 0:
        sd = v.std(ddof=1)

        st.markdown(
            f"""
**Likert 統計**
- 全部樣本數 N = {total_n}
- 有效樣本數（排除不知道） n = {valid_n}
- 不知道 / 缺失 = {unknown_n}
- 平均值 Mean = **{v.mean():.2f}**
- 標準差 SD = **{sd:.2f}**
"""
        )

#st.write("schema type:", schema.get(q))
#st.write("norm col exists:", q + "_norm" in df.columns)
if q + "_norm" in df.columns:
    st.write(df[q + "_norm"].explode().value_counts().head(30))

y = "percent" if as_percent else "count"

if group == "(不分組)":
    fig = px.bar(result, x=result.columns[0], y=y)
    st.plotly_chart(fig, width="stretch")
    st.dataframe(result)
else:
    # 分組時 result 欄位順序是: [group, q_use, count, (percent)]
    x_col = result.columns[1]   # 這就是 q_use
    fig = px.bar(result, x=x_col, y=y, color=group, barmode="group")
    st.plotly_chart(fig, width="stretch")
    st.dataframe(result)






