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
    # 1) 讀檔
    df = pd.read_excel(path, sheet_name=sheet)

    # 2) 去掉全空欄
    df = df.dropna(axis=1, how="all")

    # 3) 統一把字串欄位做 strip（避免前後空白造成類別炸裂）
    #    注意：不要把數字欄硬轉字串，所以這裡只對 object/string 欄位處理
    obj_cols = df.select_dtypes(include=["object", "string"]).columns
    df[obj_cols] = df[obj_cols].apply(lambda s: s.astype("string").str.strip())

    auto_cols = auto_detect_likert_1to5_unknown6_cols(df)

    for col in auto_cols:
        parsed = df[col].apply(parse_likert_1to5_unknown6)
        df[col + "__num"] = parsed.apply(lambda t: t[0])  # 1~5 或 NA（給平均/SD）
        df[col + "__cat"] = parsed.apply(lambda t: t[1])  # '1'..'5' 或 '不知道'（給分布圖）

    # 4) 建 schema：優先吃 NORMALIZE_MAP 的 type（避免欄名沒寫「可複選」而判錯）
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

    # 5) 套用 normalization rules
    for col, rule in NORMALIZE_MAP.items():
        if col not in df.columns:
            continue

        # 以 rule.type 為準（沒有才退回 schema）
        qtype = rule.get("type") or schema.get(col, {}).get("type", "single")

        if qtype == "multiselect":
            df[col + "_norm"] = df[col].apply(lambda x: normalize_multiselect_cell(x, rule))
        else:
            df[col] = normalize_column_single(df[col], rule)

    #df = add_college_column(df, class_col="班級")

    return df, schema

NUM_PREFIX = re.compile(r"^\s*(\d+)\s*(.*)$", re.I)

def auto_detect_likert_1to5_unknown6_cols(df: pd.DataFrame, min_15_unique: int = 3) -> list[str]:
    """
    自動找出像「1~5 + 6不知道/unknown」這種欄位。
    - 只掃描 object/string 欄
    - 需要：1~5 至少出現 min_15_unique 個不同碼
    - 且出現 6，並且尾巴包含 不知道/unknown/don't know 之類
    """
    cols = []
    obj_cols = df.select_dtypes(include=["object", "string"]).columns

    for col in obj_cols:
        s = df[col].dropna().astype(str).str.strip()
        if s.empty:
            continue

        codes = set()
        has_unknown6 = False

        for v in s.unique():
            m = NUM_PREFIX.match(v)
            if m:
                code = int(m.group(1))
                tail = (m.group(2) or "").strip().lower()
                if 1 <= code <= 5:
                    codes.add(code)
                elif code == 6:
                    # 只有 6 且文字像 unknown/不知道 才算
                    if ("unknown" in tail) or ("don't know" in tail) or ("dont know" in tail) or ("不知" in v) or ("unknown" in v):
                        has_unknown6 = True
            else:
                # 有些資料可能直接填 unknown / 不知道（沒有數字）
                low = v.lower()
                if low in {"unknown", "don't know", "dont know"} or ("不知" in v) or ("unknown" in v):
                    has_unknown6 = True

        if (len(codes) >= min_15_unique) and has_unknown6:
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
st.title("115學年度大一新生學習適應性分析")

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

# ---- 顯示 Likert 平均與標準差（排除 6 不知道）----
num_col = q_base + "__num"
if num_col in df.columns:
    v = df[num_col].dropna()
    if not v.empty:
        st.markdown(
            f"""
**Likert 統計（排除「不知道」）**  
- 有效樣本數 n = {len(v)}  
- 平均值 Mean = **{v.mean():.2f}**  
- 標準差 SD = **{v.std(ddof=1):.2f}**
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






