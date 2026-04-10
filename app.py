import re
from pathlib import Path
import pandas as pd
import streamlit as st
import plotly.express as px
from normalization import (
    COLLEGE_ORDER,
    PREFIX_ORDER,
    infer_question_type,
    normalize_column_single,
    normalize_multiselect_cell,
    NORMALIZE_MAP,
    parse_likert_1to5_unknown6,
    add_college_column,
    get_class_info,
)

FILE_PATH = "114學年度填答總表_含班級學號12102025_去識別化.xlsx"
SHEET = "1210-已填總表"
HELP_DOC_PATH = Path(__file__).with_name("說明文件.txt")

# ---------- Data loading ----------
@st.cache_data
def load_data(path: str = FILE_PATH, sheet: str = SHEET):
    df = pd.read_excel(path, sheet_name=sheet)
    df = df.dropna(axis=1, how="all")

    obj_cols = df.select_dtypes(include=["object", "string"]).columns
    df[obj_cols] = df[obj_cols].apply(lambda s: s.astype("string").str.strip())

    auto_cols = auto_detect_likert_1to5_cols(df)

    # 建立 Likert 衍生欄位
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

    # 處理 normalization
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

        # 條件1：至少有幾種 Likert 分數
        cond1 = len(codes) >= min_15_unique

        # 條件2：大部分值是 1~5（避免誤判）
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

def summarize(
    df: pd.DataFrame,
    schema: dict,
    q: str,
    group: str | None,
    as_percent: bool,
    pct_mode: str | None = None,
) -> pd.DataFrame:
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
            if pct_mode == "全體百分比":
                total_resp = d.index.nunique() if is_ms else len(d)
                out["percent"] = out["count"] / max(total_resp, 1) * 100
            elif is_ms:
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
                # 不分組：分母用作答人數
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


def apply_normalized_order(result: pd.DataFrame, col: str, college_order, class_order=None):
    if col not in result.columns:
        return result

    if col == "學院":
        order = college_order
    elif col == "班級":
        if class_order is None:
            unique_classes = result[col].dropna().astype(str).unique()
            class_order = sorted(unique_classes, key=lambda c: parse_class_key(c, college_order))
        order = class_order
    else:
        return result

    result[col] = pd.Categorical(result[col].astype(str), categories=order, ordered=True)
    sort_cols = [col] + [c for c in result.columns if c != col]
    result = result.sort_values(sort_cols)
    return result


def parse_class_key(class_name: str, college_order=None):
    text = str(class_name or "").strip()
    if not text:
        return (len(college_order) if college_order is not None else 999, len(PREFIX_ORDER), "", 0, "")

    info = get_class_info(text)
    college = info["college"]
    prefix = info["prefix"]
    college_rank = college_order.index(college) if college_order and college in college_order else (len(college_order) if college_order else 999)
    prefix_rank = PREFIX_ORDER.index(prefix) if prefix in PREFIX_ORDER else len(PREFIX_ORDER)

    m = re.search(r'([一二三四1234])(?:年級)?\s*([A-Za-z])(?:班)?$', text)
    if m:
        year_str = m.group(1)
        class_str = m.group(2)
        year_num = {'一': 1, '二': 2, '三': 3, '四': 4, '1': 1, '2': 2, '3': 3, '4': 4}.get(year_str, 0)
        return (college_rank, prefix_rank, prefix, year_num, class_str)

    # 無法解析年級/班別時先依校院、再字串排序
    return (college_rank, prefix_rank, prefix, 0, text)


def get_percent_column_label(pct_mode: str | None, group_label: str) -> str:
    if pct_mode == "全體：圖示全體=100%":
        return "百分比（全體=100%）"
    if group_label != "(不分組)":
        return "百分比（各組=100%）"
    return "百分比"


def normalize_display_table(df: pd.DataFrame, percent_col_label: str = "百分比") -> pd.DataFrame:
    out = df.copy()
    if "option" in out.columns:
        out = out.rename(columns={"option": "選項"})
    if "count" in out.columns:
        out = out.rename(columns={"count": "人數"})
    if "percent" in out.columns:
        out = out.rename(columns={"percent": percent_col_label})
        out[percent_col_label] = out[percent_col_label].astype(float).round(2).map(lambda x: f"{x:.2f}%")
    return out


def show_table(df: pd.DataFrame, percent_col_label: str = "百分比", **kwargs):
    st.dataframe(normalize_display_table(df, percent_col_label=percent_col_label), hide_index=True, width="stretch", **kwargs)


def show_table_caption(text: str):
    st.markdown(
        f"<div style='font-size: 1rem; color: #111111; margin: 0.25rem 0 0.5rem 0;'>{text}</div>",
        unsafe_allow_html=True,
    )


def is_groupable_column(df: pd.DataFrame, schema: dict, col: str) -> bool:
    if col not in df.columns:
        return False

    excluded = {"已填人", "開始時間", "完成時間", "上次修改時間", "填答時間", "學號", "姓名", "Email", "班級前綴"}
    if col in excluded:
        return False

    if schema.get(col, {}).get("type") == "multiselect":
        return False

    s = df[col].dropna().astype(str).str.strip()
    s = s[s != ""]
    if s.empty:
        return False

    unique_count = s.nunique()
    sample_size = len(s)

    # 避免把高基數欄位（近似自由填答/識別欄）塞進分組清單。
    if unique_count <= min(30, max(12, int(sample_size * 0.2))):
        return True

    return col in {"學院", "班級"}


def build_population_text(selected_colleges: list[str], selected_classes: list[str]) -> str:
    parts = []
    if selected_colleges:
        parts.append("、".join(selected_colleges))
    if selected_classes:
        parts.append("、".join(selected_classes))
    return "；".join(parts)


def build_table_caption(
    question_label: str,
    group_label: str,
    selected_colleges: list[str],
    selected_classes: list[str],
) -> str:
    population_text = build_population_text(selected_colleges, selected_classes)
    grouped = group_label != "(不分組)"
    filtered = bool(selected_colleges or selected_classes)

    if grouped and filtered:
        return f"{question_label}依{group_label}篩選在{population_text}統計"
    if grouped:
        return f"{question_label}依{group_label}的全校統計"
    if filtered:
        return f"{question_label}在{population_text}的統計"
    return f"{question_label}的全校統計"


def load_help_document(path: Path = HELP_DOC_PATH) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except FileNotFoundError:
        return "尚未建立說明文件。"


# ---------- UI ----------
st.set_page_config(page_title="115學年度大一新生學習適應性分析", layout="wide")
st.title("115學年度大一新生學習適應性分析")

if "show_help_doc" not in st.session_state:
    st.session_state.show_help_doc = False

if st.session_state.show_help_doc:
    st.subheader("說明文件")
    st.markdown(load_help_document())
    st.divider()

df, schema = load_data(FILE_PATH)

df = add_college_column(df)

# 根據 normalization.py 的註解順序固定學院/學系前綴排序
college_order = COLLEGE_ORDER

all_questions = [
    c for c in df.columns
    if c not in ["已填人"]
    and not c.endswith("__num")
    and not c.endswith("__cat")
    and not c.endswith("_norm")
]

default_groups = ["(不分組)"]
preferred_group_order = [
    "性別",
    "身分別",
    "學院",
    "班級",
    "原畢業學校之類型",
    "原畢業學校所在地區",
]
candidate_groups = [c for c in preferred_group_order if c in df.columns and is_groupable_column(df, schema, c)]

extra_group_candidates = [
    c for c in all_questions
    if c not in candidate_groups and is_groupable_column(df, schema, c)
]

group_options = default_groups + candidate_groups + extra_group_candidates

with st.sidebar:
    st.header("分析設定")
    q = st.selectbox("問卷題目（圖表類別）", all_questions)
    available_group_options = [opt for opt in group_options if opt == "(不分組)" or opt != q]
    group = st.selectbox("分組比較（群組標籤）", available_group_options)

    st.divider()

    # 母體篩選：學院、班級 可複選
    population_attrs = st.multiselect(
        "學院、班級篩選（可複選交叉比對或留空表示全校）", 
        ["學院", "班級"],
        default=[],
        placeholder="不篩選(全校)",
    )

    selected_colleges = []
    selected_classes = []
    if "學院" in population_attrs and "學院" in df.columns:
        college_values = [x for x in college_order if x in df["學院"].dropna().astype(str).unique()]
        extras = [x for x in sorted(df["學院"].dropna().astype(str).unique()) if x not in college_values]
        selected_colleges = st.multiselect(
            "選取學院（可多選）",
            college_values + extras,
            default=[],
            placeholder="不篩選(全校)",
        )
    if "班級" in population_attrs and "班級" in df.columns:
        class_values = df["班級"].dropna().astype(str).unique()
        class_ordered = sorted(class_values, key=lambda c: parse_class_key(c, college_order))
        selected_classes = st.multiselect(
            "選取班級（可多選）",
            class_ordered,
            default=[],
            placeholder="不篩選(全校)",
        )

    st.divider()

    as_percent = st.checkbox("顯示百分比 (%)", value=True)

    if as_percent:
        pct_mode = st.radio(
            "百分比母體",
            ["全體：圖示全體=100%", "分組百分比：各組各自總和=100%"],
            index=1,
        )
    else:
        pct_mode = None

    st.divider()
    if st.button("說明文件", use_container_width=True):
        st.session_state.show_help_doc = not st.session_state.show_help_doc
    st.caption(f"文件：{HELP_DOC_PATH.name}")

# Apply population filter (母體)：符合任一選項
mask = pd.Series(True, index=df.index)
if population_attrs:
    if selected_colleges or selected_classes:
        mask = pd.Series(False, index=df.index)
        if selected_colleges:
            mask |= df["學院"].isin(selected_colleges)
        if selected_classes:
            mask |= df["班級"].isin(selected_classes)
    else:
        # 已選欄位但未選值：顯示空結果。
        mask = pd.Series(False, index=df.index)

filtered_df = df[mask]
if filtered_df.empty:
    st.warning("母體篩選後無資料，請調整學院 / 班級選擇。")

# 若為學院/班級相關問題，先套用 normalization.py 提供的預設排序
grouped = (group != "(不分組)")
result = summarize(
    filtered_df,
    schema,
    q=q,
    group=None if group == "(不分組)" else group,
    as_percent=as_percent,
    pct_mode=pct_mode,
)

if grouped:
    x_col = result.columns[1]
    group_col = result.columns[0]
else:
    x_col = result.columns[0]
    group_col = None

if x_col in ["學院", "班級"]:
    result = apply_normalized_order(result, x_col, college_order)

if grouped and group_col in ["學院", "班級"]:
    result = apply_normalized_order(result, group_col, college_order)

q_base = q.replace("__cat", "").replace("__num", "")
table_caption = build_table_caption(q_base, group, selected_colleges, selected_classes)
percent_col_label = get_percent_column_label(pct_mode, group)
st.divider()

# ---- 顯示 Likert 平均與標準差（只要有 1~5 就顯示）----
num_col = f"{q_base}__num"
if num_col in df.columns:
    raw_all = pd.to_numeric(df[num_col], errors="coerce")
    v_all = raw_all.dropna()

    total_n_all = len(raw_all)
    valid_n_all = len(v_all)
    unknown_n_all = total_n_all - valid_n_all

    raw_filtered = pd.Series(dtype="float")
    v_filtered = pd.Series(dtype="float")
    total_n_filtered = valid_n_filtered = unknown_n_filtered = 0
    if not filtered_df.empty:
        raw_filtered = pd.to_numeric(filtered_df[num_col], errors="coerce")
        v_filtered = raw_filtered.dropna()
        total_n_filtered = len(raw_filtered)
        valid_n_filtered = len(v_filtered)
        unknown_n_filtered = total_n_filtered - valid_n_filtered

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**全校 Likert 統計（目前的為全校）**")
        if valid_n_all > 0:
            sd_all = v_all.std(ddof=1)
            st.markdown(
                f"""
- 全部樣本數 N = {total_n_all}
- 有效樣本數（排除不知道） n = {valid_n_all}
- 不知道 / 缺失 = {unknown_n_all}
- 平均值 Mean = **{v_all.mean():.2f}**
- 標準差 SD = **{sd_all:.2f}**
"""
            )
        else:
            st.write("沒有有效 Likert 數值資料（全校）。")

    filtered_active = bool(population_attrs and (selected_colleges or selected_classes))
    with col2:
        if filtered_active:
            st.markdown("**篩選後 Likert 統計**")
            if valid_n_filtered > 0:
                sd_filtered = v_filtered.std(ddof=1)
                st.markdown(
                    f"""
- 全部樣本數 N = {total_n_filtered}
- 有效樣本數（排除不知道） n = {valid_n_filtered}
- 不知道 / 缺失 = {unknown_n_filtered}
- 平均值 Mean = **{v_filtered.mean():.2f}**
- 標準差 SD = **{sd_filtered:.2f}**
"""
                )
            else:
                st.write("沒有有效 Likert 數值資料（篩選後）。")
        else:
            st.write("未選擇母體篩選或未有選值，篩選後統計與全校相同。")

y = "percent" if as_percent else "count"
y_axis_label = "百分比(%)" if as_percent else "人數"

category_orders = {}
if x_col in ["學院", "班級"]:
    if x_col == "學院":
        category_orders[x_col] = college_order
    else:  # 班級
        unique_x = result[x_col].dropna().astype(str).unique()
        category_orders[x_col] = sorted(unique_x, key=lambda c: parse_class_key(c, college_order))

if grouped and group in ["學院", "班級"]:
    if group == "學院":
        category_orders[group] = college_order
    else:  # 班級
        unique_group = result[group].dropna().astype(str).unique()
        category_orders[group] = sorted(unique_group, key=lambda c: parse_class_key(c, college_order))

if group == "(不分組)":
    fig = px.bar(result, x=x_col, y=y, category_orders=category_orders if category_orders else None)
    fig.update_yaxes(title=y_axis_label)
    st.plotly_chart(fig, width="stretch")
    show_table_caption(table_caption)
    show_table(result, percent_col_label=percent_col_label)
else:
    # 分組時 result 欄位順序是: [group, q_use, count, (percent)]
    compact_mode = (x_col == "班級" and group == "學院")
    fig = px.bar(
        result,
        x=x_col,
        y=y,
        color=group,
        barmode="stack" if compact_mode else "group",
        category_orders=category_orders if category_orders else None,
    )
    if compact_mode:
        fig.update_layout(bargap=0.12, bargroupgap=0)
    fig.update_yaxes(title=y_axis_label)
    st.plotly_chart(fig, width="stretch")
    show_table_caption(table_caption)
    show_table(result, percent_col_label=percent_col_label)






