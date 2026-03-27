import pandas as pd
import re

# 依照「十八學群及其學類對照表」建立（可逐步補充同義詞/系名）
#靜宜大學順序
#外語學院	英文系
#外語學院	日文系
#外語學院	西文系
#人社院	中文系
#人社院	社工系
#人社院	台文系
#人社院	法律系
#人社院	大傳系
#人社院	生態系
#人社院	法律原專
#人社院	社工原專
#理學院	財工系
#理學院	應化系
#理學院	食營系
#理學院	化科系
#理學院	永續環境與智慧科技學士學位學程
#管理學院	行銷與數位經營學系
#管理學院	國企系
#管理學院	會計系
#管理學院	觀光系
#管理學院	財金系
#資訊學院	資管系
#資訊學院	資工系
#資訊學院	人工智慧系
#資訊學院	資工系
#資訊學院	資科系
#國際學院	國際資訊學士學位學程
#國際學院	寰宇外語教育學士學位學程
#國際學院	寰宇管理學士學位學程

GROUP_TO_CLASSES = {
    "資訊學群": ["資訊工程", "數據統計", "電機工程", "光電工程", "電子工程", "通訊工程", "生物資訊", "資訊傳播", "圖書資訊", "數位學習", "資訊管理", "電子商務", "媒體設計"],
    "工程學群": ["機械工程", "航空工程", "土木工程", "水利工程", "化學工程", "材料工程", "工程科學", "環境工程", "建築", "運輸物流", "科技管理", "工程不分系", "電資不分系"],
    "數理化學群": ["數學", "化學", "物理", "自然科學", "生化", "財金統計", "理學不分系", "數據統計"],
    "醫藥衛生學群": ["醫學", "公共衛生", "牙醫", "物理治療", "職能治療", "護理", "醫學檢驗", "影像放射", "藥學", "食品營養", "呼吸治療", "健康照護", "化妝品", "職業安全", "視光", "語療聽力", "醫務管理", "獸醫"],
    "生命科學學群": ["生命科學", "生物科技", "生態", "食品生技"],
    "生物資源學群": ["植物保護", "農藝", "動物科學", "園藝", "森林", "海洋資源"],
    "地球環境學群": ["地球科學", "地理", "海洋科學", "大氣科學", "防災", "史地"],
    "建築設計學群": ["都市計畫", "空間設計", "工業設計", "工藝", "商業設計", "服裝設計", "藝術設計", "藝術不分系"],
    "藝術學群": ["美術", "音樂", "表演藝術", "舞蹈"],
    "社會心理學群": ["心理", "社會學", "社會工作", "人類民族", "兒童家庭", "輔導諮商", "勞工關係"],
    "大眾傳播學群": ["大眾傳播", "廣電電影", "新聞", "廣告公關"],
    "外語學群": ["英語文", "歐語文", "日語文", "東方語文", "英語教育", "華語文教育"],
    "文史哲學群": ["中國語文", "歷史", "哲學", "台灣語文", "宗教", "文化產業"],
    "教育學群": ["教育", "特殊教育", "幼兒教育", "成人教育", "社科教育", "科技教育", "數學教育", "英語教育", "華語文教育"],
    "法政學群": ["法律", "財經法律", "政治", "行政管理", "犯罪防治", "土地資產"],
    "管理學群": ["企業管理", "行銷經營", "國際企業", "觀光事業", "運動管理", "餐旅管理", "休閒管理", "商管不分系", "醫務管理", "科技管理"],
    "財經學群": ["會計", "財務金融", "財稅", "保險", "經濟"],
    "遊憩運動學群": ["體育", "運動保健", "觀光事業", "運動管理", "餐旅管理", "休閒管理"],
    # 這裡把「不分系」單獨當一類（PDF 有說不分系不屬於學群學類架構）。:contentReference[oaicite:2]{index=2}
    "不分系": ["學院不分系", "不分系", "工程不分系", "藝術不分系", "電資不分系", "商管不分系", "理學不分系"],
}

PREFIX_TO_COLLEGE = {
    # 外語學院
    "英": "外語學院",
    "日": "外語學院",
    "西": "外語學院",

    # 人文暨社會科學學院
    "中": "人文暨社會科學學院",
    "大傳": "人文暨社會科學學院",          # 大眾傳播（你資料用大傳）
    "台文": "人文暨社會科學學院",          # 臺灣文學
    "法律": "人文暨社會科學學院",
    "社工": "人文暨社會科學學院",
    "社工原專": "人文暨社會科學學院",
    "犯防": "人文暨社會科學學院",
    "犯防原專": "人文暨社會科學學院",
    "生態": "人文暨社會科學學院",          # 依你截圖：生態也出現在法律/社工附近，先放人社；若你校內其實屬理學院再改這行

    # 理學院
    "食營": "理學院",
    "應化": "理學院",
    "化科": "理學院",                      # 你資料有 化科一A/化科一B
    "財工": "理學院",                      # 財務工程（你資料用財工）
    "永續智慧": "理學院",                  # 你資料有 永續智慧一A（若校內其實歸國際/其他學院，再改）

    # 管理學院
    "國企": "管理學院",
    "行銷與數位經營": "管理學院",
    "會計": "管理學院",
    "財金": "管理學院",
    "觀光": "管理學院",
    "經管進": "管理學院",                  # 你資料有 經管進一A（推測：經營管理進修？）

    # 資訊學院
    "資管": "資訊學院",
    "資工": "資訊學院",
    "資科": "資訊學院",
    "人工智慧": "資訊學院",
    "智慧媒體學程": "資訊學院",
    "晶片設計": "資訊學院",                # 你資料有 晶片設計一A（若校內歸工學/理學請改）
    "國際資訊學士學位學程": "資訊學院",

    # 國際學院
    "寰宇管理學程": "國際學院",
    "寰宇外語教育": "國際學院",
}

_PREFIX_RE = re.compile(r"^(?P<prefix>.+?)(?=[一二三四五六七八九十])")

def extract_class_prefix(class_name: str) -> str | None:
    if pd.isna(class_name):
        return None
    s = str(class_name).strip()
    m = _PREFIX_RE.search(s)
    return m.group("prefix") if m else None

def infer_college_from_class(class_name: str) -> str | None:
    prefix = extract_class_prefix(class_name)
    if not prefix:
        return None

    # 先做完全命中（長字優先）
    for k in sorted(PREFIX_TO_COLLEGE.keys(), key=len, reverse=True):
        if prefix == k:
            return PREFIX_TO_COLLEGE[k]

    # 再做開頭命中（救少數異常）
    for k in sorted(PREFIX_TO_COLLEGE.keys(), key=len, reverse=True):
        if prefix.startswith(k):
            return PREFIX_TO_COLLEGE[k]

    return None

def add_college_column(df: pd.DataFrame, class_col: str = "班級",
                       out_col: str = "學院", prefix_col: str = "班級前綴") -> pd.DataFrame:
    if class_col not in df.columns:
        return df
    df[prefix_col] = df[class_col].apply(extract_class_prefix)
    df[out_col] = df[class_col].apply(infer_college_from_class)
    return df

# 反向：keyword(學類/常見系名) -> 學群
KEYWORD_TO_GROUP = {}
for g, classes in GROUP_TO_CLASSES.items():
    for c in classes:
        KEYWORD_TO_GROUP[c] = g

# 你可以在這裡補「常見系名縮寫/別名」(非常有用)
ALIAS = {
    "資工": "資訊工程",
    "電機": "電機工程",
    "機械": "機械工程",
    "土木": "土木工程",
    "化工": "化學工程",
    "材科": "材料工程",
    "護理系": "護理",
    "藥學系": "藥學",
    "企管": "企業管理",
    "財金": "財務金融",
    "會計系": "會計",
    "觀光系": "觀光事業",
    "餐旅": "餐旅管理",
}
def normalize_q16_academic_group(value: object) -> str:
    """把第16題（學群/學類/系名/其他）正規化到：18學群 + 不分系 + 其他(仍判不到)"""
    if pd.isna(value):
        return ""

    s = str(value).strip()
    if not s:
        return ""

    s_low = s.lower()

    # 1) 若學生直接填「某某學群」，直接回傳
    for g in GROUP_TO_CLASSES.keys():
        if g in s:
            return g

    # 2) 先吃 alias（系名縮寫 → 學類 → 學群）
    for k, v in ALIAS.items():
        if k in s:
            s = s.replace(k, v)

    # 3) 用「包含」方式找學類/關鍵詞（把“其他：OO系/OO學類”也吃掉）
    for kw, g in KEYWORD_TO_GROUP.items():
        if kw and kw in s:
            return g

    # 4) 常見的「其他」但其實沒填任何線索
    if "其他" in s or "other" in s_low:
        return "其他"

    return "其他"

MULTI_NAME_RE = re.compile(r"(可複選|複選|多選)")
FULLWIDTH_SEMI = "；"
SPLIT_RE = re.compile(r"\s*[；;]\s*")

def infer_question_type(col_name: str, series: pd.Series, sample_n: int = 200, semi_ratio: float = 0.15) -> str:
    if MULTI_NAME_RE.search(str(col_name)):
        return "multiselect"

    s = series.dropna().astype(str)
    if s.empty:
        return "single"
    if len(s) > sample_n:
        s = s.sample(sample_n, random_state=0)

    ratio = s.str.contains(FULLWIDTH_SEMI, regex=False).mean()
    return "multiselect" if ratio >= semi_ratio else "single"

def normalize_column_single(series: pd.Series, rule: dict) -> pd.Series:
    """single 題：contains 先、exact 後；mapping key 假設已 lower。"""
    def _norm(x):
        if pd.isna(x):
            return pd.NA
        key = str(x).strip().lower()

        for pat, target in rule.get("contains", {}).items():
            if pat in key:
                return target

        return rule.get("exact", {}).get(key, str(x).strip())

    if rule.get("custom") == "q16_academic_group":
        return series.apply(normalize_q16_academic_group)

    # 否則維持你原本 exact/contains 的邏輯...
    ...

    return series.apply(_norm)

def split_multiselect_cell(cell) -> list[str]:
    if pd.isna(cell):
        return []
    raw = str(cell).strip()
    if not raw:
        return []
    parts = [p.strip() for p in SPLIT_RE.split(raw)]
    return [p for p in parts if p]

def normalize_token(token: str, rule: dict) -> str:
    key = str(token).strip().lower()

    for pat, target in rule.get("contains", {}).items():
        if pat in key:
            return target

#    if key not in rule.get("exact", {}):
#        print("[UNMAPPED TOKEN]", token)

    return rule.get("exact", {}).get(key, token.strip())


def normalize_multiselect_cell(cell, rule: dict) -> list[str]:
    parts = split_multiselect_cell(cell)
    normalized = [normalize_token(p, rule) for p in parts]
    normalized = [p for p in normalized if p]

    # 去重保序
    seen = set()
    out = []
    for p in normalized:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out


NORMALIZE_MAP = {
    "1性別": {
        "exact": {"male": "男性",
        "m": "男性",
        "男性": "男性",
        "男": "男性",
        "female": "女性",
        "f": "女性",
        "女性": "女性",
        "女": "女性",
        }
    },
    "2身分別": {
        "contains": {
            "foreign student": "外籍生",
            "international student": "外籍生",
            "外籍": "外籍生"
        },
        "exact": {
            "本地生": "本國生",
            "local student": "本國生",
            "僑生": "僑生",
            "本國生": "本國生",
        }
    },
    "3家庭經濟狀況（可複選）": {
        "type": "multiselect",
        "contains": {
            # 低/中低收入
            "低收入": "低收入/中低收入",
            "中低收入": "低收入/中低收入",
            "low-income": "中低收入戶學生",
            "low income": "中低收入戶學生",
            "middle-income": "中低收入戶學生",
            "hardship": "清寒學生",
            "children/grandchildren": "清寒學生",
            # 弱勢類別
            "身心障礙": "身心障礙",
            "原住民": "原住民",
            "特殊境遇": "特殊境遇家庭",
            "懷孕": "懷孕/育兒(3歲以下)",
            "分娩": "懷孕/育兒(3歲以下)",
            "撫育3歲以下": "懷孕/育兒(3歲以下)",
            "家庭突遭變故": "家庭突遭變故",
            "弱勢助學金": "弱勢助學金",
            "學雜費減免資格但獲教育部弱勢助學金": "弱勢助學金",  # 可選，保險

            # 以上皆非 / 其他（如果資料會出現）
            "以上皆非": "以上皆非",
            "none of the above": "以上皆非",
            "others": "其他",
            "other": "其他",
            "其他": "其他",
        },
        "exact": {
            # 讓已標準化的值保持不變（可有可無，但很乾淨）
            "低收入/中低收入": "低收入/中低收入",
            "身心障礙": "身心障礙",
            "原住民": "原住民",
            "特殊境遇家庭": "特殊境遇家庭",
            "懷孕/育兒(3歲以下)": "懷孕/育兒(3歲以下)",
            "家庭突遭變故": "家庭突遭變故",
            "弱勢助學金": "弱勢助學金",
            "以上皆非": "以上皆非",
            "其他": "其他",
        },
        "unknown_to_other": False,
    },
    "4家庭文化背景（可複選）": {
        "type": "multiselect",
        "contains": {
            # 以上皆非
            "以上皆非": "以上皆非",
            "none of the above": "以上皆非",

            # 三代家庭 / 家中第一代大專（中文長句 + 英文版本）
            "三代家庭": "三代家庭/家中第一代大專",
            "原生家庭第1位上大專": "三代家庭/家中第一代大專",
            "第1位上大專": "三代家庭/家中第一代大專",
            "first one to study": "三代家庭/家中第一代大專",
            "first one to study at university": "三代家庭/家中第一代大專",
            "the first one to study at university": "三代家庭/家中第一代大專",
            "the first one to study at university/college": "三代家庭/家中第一代大專",

            # 新住民
            "新住民": "新住民/新住民二代",
            "new immigrant": "新住民/新住民二代",

            # 其他（含自由填答）
            "其他 /": "其他",
            "other /": "其他",
            "others /": "其他",
        },
        "exact": {
            # 讓已標準化的值保持不變（可有可無，但很乾淨）
            "以上皆非": "以上皆非",
            "三代家庭/家中第一代大專": "三代家庭/家中第一代大專",
            "新住民/新住民二代": "新住民/新住民二代",
            "其他": "其他",
        },
        # 建議打開：少數沒命中的長句/奇怪輸入，直接收斂到「其他」
        "unknown_to_other": True,
    },
    "5原畢業學校之類型": {
        "type": "single",
        "contains": {
            "普通": "普通高中",
            "general high": "普通高中",
            "high school": "普通高中",

            "高職": "高職/技術型高中",
            "技術型": "高職/技術型高中",
            "vocational": "高職/技術型高中",

            "綜合高中": "綜合高中",
            "home": "其他",
            "自主學習": "其他",
            "university": "其他",
            "other": "其他",
        },
        "exact": {
            "普通高中": "普通高中",
            "高職/技術型高中": "高職/技術型高中",
            "綜合高中": "綜合高中",
            "其他": "其他",
        },
        "unknown_to_other": True,
    },
    "6原畢業學校所在地區": {
        "type": "single",
        "contains": {
            "foreign": "國外",
            "國外": "國外",
            "taichung": "臺中市",
            "taipei": "臺北市",
            "kaohsiung": "高雄市",
            "taoyuan": "桃園市",
            "tainan": "臺南市",
            "hsinchu": "新竹市",
            "pingtung": "屏東縣",
            "other": "其他",
        },
        "exact": {
            "國外": "國外",
        },
        "unknown_to_other": False,
    },
    "7大學「學費」主要來源（可複選）": {
        "type": "multiselect",
        "contains": {
            "家庭提供": "家庭提供",
            "助學貸款": "助學貸款",
            "學雜費減免": "學雜費減免",
            "打工": "打工或工讀所得",
            "工讀": "打工或工讀所得",
            "獎助學金": "校內外獎助學金",
            "financial aid": "校內外獎助學金",
            "scholarship": "校內外獎助學金",
            "part-time": "打工或工讀所得",
            "part time": "打工或工讀所得",
            "contributions/remittance": "家庭提供",
            "other /": "其他",
            "其他 /": "其他",
            "others": "其他",
        },
        "exact": {
            "家庭提供": "家庭提供",
            "助學貸款": "助學貸款",
            "學雜費減免": "學雜費減免",
            "打工或工讀所得": "打工或工讀所得",
            "校內外獎助學金": "校內外獎助學金",
            "其他": "其他",
        },
    "unknown_to_other": False,
    },
    "8學習及生活費（書籍、住宿、交通、伙食等開銷）主要來源（可複選）":{
        "type": "multiselect",
        "contains": {
            # 家庭提供
            "家庭提供": "家庭提供",
            "contributions/remittance from family": "家庭提供",
            "remittance from family": "家庭提供",
            "from family": "家庭提供",

            # 打工 / 工讀
            "打工": "打工或工讀所得",
            "工讀": "打工或工讀所得",
            "earnings from part-time job": "打工或工讀所得",
            "part-time job": "打工或工讀所得",

            # 獎助學金
            "獎助學金": "校內外獎助學金",
            "scholarship": "校內外獎助學金",

            # 助學貸款
            "助學貸款": "助學貸款",
            "student loan": "助學貸款",
            "loan": "助學貸款",

            # 政府補助
            "政府補助": "政府補助",
            "financial aid": "政府補助",
            "financial aids": "政府補助",

            # 其他（含自由填答）
            "其他 /": "其他",
            "other /": "其他",
            "others /": "其他",
            "其他": "其他",
        },
        "exact": {
            "家庭提供": "家庭提供",
            "打工或工讀所得": "打工或工讀所得",
            "校內外獎助學金": "校內外獎助學金",
            "助學貸款": "助學貸款",
            "政府補助": "政府補助",
            "其他": "其他",
        },
        # 少數沒命中就收斂到「其他」
        "unknown_to_other": True,
    },
    "9我的入學管道": {
        "type": "single",
        "contains": {
            "personal application": "大學個人申請",
            "大學個人申請": "大學個人申請",

            "繁星": "大學繁星",
            "大學繁星": "大學繁星",

            "考試分發": "大學考試分發",
            "大學考試分發": "大學考試分發",

            "四技二專登記分發": "四技二專登記分發",
            "joint enrolment and allocation": "四技二專登記分發",

            "四技二專甄選": "四技二專甄選",
            "selection by recommendation": "四技二專甄選",

            "技優": "四技二專技優",
            "四技二專技優": "四技二專技優",

            "independent recruitment": "獨立招生",
            "獨立招生": "獨立招生",
            "特殊選才": "獨立招生",

            "overseas": "海外聯合招生",
            "海外聯合招生": "海外聯合招生",

            "身心障礙": "身心障礙甄試",
            "運動": "運動績優甄試",

            "international student": "外籍生/國際生",
            "agent": "外籍生/國際生",

            "轉系": "其他",
            "轉學": "其他",
            "other": "其他",
            "其他 /": "其他",
        },
        "exact": {
            "大學個人申請": "大學個人申請",
            "大學繁星": "大學繁星",
            "大學考試分發": "大學考試分發",
            "四技二專登記分發": "四技二專登記分發",
            "四技二專甄選": "四技二專甄選",
            "四技二專技優": "四技二專技優",
            "獨立招生": "獨立招生",
            "海外聯合招生": "海外聯合招生",
            "身心障礙甄試": "身心障礙甄試",
            "運動績優甄試": "運動績優甄試",
            "外籍生/國際生": "外籍生/國際生",
            "其他": "其他",
        },
        "unknown_to_other": True,
    },
    "10得知本校最主要的管道（可複選）": {
        "type": "multiselect",
        "contains": {
            # 網路
            "網路": "網路資訊",
            "the internet": "網路資訊",
            "internet": "網路資訊",

            # 簡章
            "報考簡章": "報考簡章",
            "brochure": "報考簡章",
            "prospectus": "報考簡章",

            # 老師/學校宣導
            "高中職/五專老師": "高中職/五專老師",
            "high school/technical school teachers": "高中職/五專老師",
            "teachers from your current school": "高中職/五專老師",

            "本校進班宣導": "本校進班宣導",
            "campus promotion": "本校進班宣導",

            "本校教師": "本校教師",

            # 同儕/家人
            "朋友、同學、學長姐": "朋友、同學、學長姐",
            "friends, classmates": "朋友、同學、學長姐",
            "senior classmates": "朋友、同學、學長姐",

            "家人": "家人",
            "family": "家人",

            # 聲望
            "學校聲望": "學校聲望",
            "school reputation": "學校聲望",

            "科系聲望": "科系聲望",
            "department reputation": "科系聲望",

            # 補習班
            "補習班": "補習班宣導",

            # 廣告（公車/看板/招牌）
            "立牌廣告": "學校立牌廣告（公車/看板/招牌）",
            "公車廣告": "學校立牌廣告（公車/看板/招牌）",
            "招牌廣告": "學校立牌廣告（公車/看板/招牌）",
            "billboard": "學校立牌廣告（公車/看板/招牌）",
            "advertisements": "學校立牌廣告（公車/看板/招牌）",

            # 其他（含自由填答）
            "其他 /": "其他",
            "other /": "其他",
            "others /": "其他",
            "其他": "其他",
        },
        "exact": {
            "網路資訊": "網路資訊",
            "報考簡章": "報考簡章",
            "高中職/五專老師": "高中職/五專老師",
            "朋友、同學、學長姐": "朋友、同學、學長姐",
            "家人": "家人",
            "學校聲望": "學校聲望",
            "科系聲望": "科系聲望",
            "本校進班宣導": "本校進班宣導",
            "本校教師": "本校教師",
            "補習班宣導": "補習班宣導",
            "學校立牌廣告（公車/看板/招牌）": "學校立牌廣告（公車/看板/招牌）",
            "其他": "其他",
        },
        # 少數奇怪輸入就收斂到「其他」，避免類別爆炸
        "unknown_to_other": True,
    },
    "11決定就讀本校最主要的原因（可複選）":{
        "type": "multiselect",
        "contains": {
            # 經濟 / 學費
            "學費": "學費/經濟考量",
            "經濟": "學費/經濟考量",
            "financial": "學費/經濟考量",
            "tuition": "學費/經濟考量",
            "cheap": "學費/經濟考量",

            # 地理位置 / 交通
            "地理": "地緣關係",
            "交通": "地緣關係",
            "距離": "地緣關係",
            "location": "地緣關係",
            "close to": "地緣關係",
            "neighborhood and location": "地緣關係",

            # 學校聲望
            "學校聲望": "學校聲望",
            "校譽": "學校聲望",
            "school reputation": "學校聲望",
            "reputation of the school": "學校聲望",
            "department reputation": "科系聲望",

            # 師資
            "師資": "師資",
            "faculty": "師資",
            "teachers": "師資",

            # 入學門檻 / 能上
            "考得上": "考試成績決定",
            "能上": "考試成績決定",
            "成績": "考試成績決定",
            "錄取": "考試成績決定",
            "admission": "考試成績決定",
            "accepted": "考試成績決定",
            "grade": "考試成績決定",
            "score": "考試成績決定",

            # 他人建議（家人/師長/同儕都併這類：先求收斂）
            "家人": "家人的影響或建議",
            "朋友": "朋友、同學、學長姐的影響或建議",
            "老師": "老師的影響或建議",
            "parents": "家人的影響或建議",
            "friends": "朋友、同學、學長姐的影響或建議",
            "advice from family": "家人的影響或建議",
            "advice from teachers": "老師的影響或建議",
            "classmates and friends": "朋友、同學、學長姐的影響或建議",
            "classmates": "朋友、同學、學長姐的影響或建議",

            # 校園環境 / 設備
            "環境": "校園環境/設備",
            "設備": "校園環境/設備",
            "campus": "校園環境/設備",
            "environment and facilities": "校園環境/設備",
            "facilities": "校園環境/設備",
            "國際化": "校園國際化程度",
            "international": "校園國際化程度",
            "internalization": "校園國際化程度", # 打錯字?
            "global": "校園國際化程度",
            "goal/objectives": "認同學校教育目標",

            # 生涯 / 職涯
            "生涯": "符合生涯發展規劃",
            "career": "符合生涯發展規劃",

            # 其他 / 以上皆非
            "以上皆非": "以上皆非",
            "none of the above": "以上皆非",
            "其他": "其他",
            "other": "其他",
            "others": "其他",
        },
        "exact": {
            "符合生涯發展規劃": "符合生涯發展規劃",
            "老師的影響或建議": "老師的影響或建議",
            "學費/經濟考量": "學費/經濟考量",
            "地緣關係": "地緣關係",
            "學校聲望": "學校聲望",
            "認同學校教育目標": "認同學校教育目標",
            "校園國際化程度": "校園國際化程度",
            "師資": "師資",
            "科系聲望": "科系聲望",
            "考試成績決定": "考試成績決定",
            "朋友、同學、學長姐的影響或建議": "朋友、同學、學長姐的影響或建議",
            "校園環境/設備": "校園環境/設備",
            "以上皆非": "以上皆非",
            "其他": "其他",
        },
        "unknown_to_other": True,
    },
    "12我選擇目前就讀「科系」的動機（可複選）": {
        "type": "multiselect",
        "contains": {
            # 興趣 / 適性
            "興趣": "興趣/適性",
            "適性": "興趣/適性",
            "喜歡": "興趣/適性",
            "interest": "興趣/適性",
            "aptitude": "興趣/適性",

            # 未來出路 / 就業
            "就業": "未來出路/就業",
            "出路": "未來出路/就業",
            "工作": "未來出路/就業",
            "薪資": "未來出路/就業",
            "career": "未來出路/就業",
            "job": "未來出路/就業",
            "employment": "未來出路/就業",

            # 科系聲望 / 評價
            "科系聲望": "科系聲望/評價",
            "系上聲望": "科系聲望/評價",
            "department reputation": "科系聲望/評價",
            "reputation": "科系聲望/評價",

            # 課程內容 / 特色
            "課程": "課程內容/特色",
            "內容": "課程內容/特色",
            "特色": "課程內容/特色",
            "curriculum": "課程內容/特色",
            "courses": "課程內容/特色",

            # 師資
            "師資": "師資",
            "老師": "師資",
            "faculty": "師資",
            "teachers": "師資",

            # 家人/同儕建議
            "家人": "朋友、同學、學長姐的影響或建議",
            "父母": "朋友、同學、學長姐的影響或建議",
            "朋友": "朋友、同學、學長姐的影響或建議",
            "同學": "朋友、同學、學長姐的影響或建議",
            "學長姐": "朋友、同學、學長姐的影響或建議",
            "建議": "朋友、同學、學長姐的影響或建議",
            "parents": "朋友、同學、學長姐的影響或建議",
            "friends": "朋友、同學、學長姐的影響或建議",
            "recommend": "朋友、同學、學長姐的影響或建議",

            # 分數/能上/錄取
            "分數": "考試成績決定",
            "考試": "考試成績決定",
            "能上": "考試成績決定",
            "錄取": "考試成績決定",
            "admission": "考試成績決定",
            "accepted": "考試成績決定",

            # 地理/交通（如果12題也有）
            "地理": "地緣關係",
            "交通": "地緣關係",
            "距離": "地緣關係",
            "location": "地緣關係",

            # 經濟/學費（如果12題也有）
            "學費": "學費/經濟考量",
            "經濟": "學費/經濟考量",
            "tuition": "學費/經濟考量",
            "financial": "學費/經濟考量",

            # 其他 / 以上皆非
            "以上皆非": "以上皆非",
            "none of the above": "以上皆非",
            "其他 /": "其他",
            "other /": "其他",
            "others /": "其他",
            "其他": "其他",
            "other": "其他",
            "others": "其他",
            
            # --- 英文翻譯併回中文（第12題） ---

            # 專業能力
            "growing professional competence": "想學習專業知能",

            # 就業 / 出路
            "employability": "未來出路/就業",

            # 原本就讀學科
            "your previous major": "受原本就讀學科影響",

            # 升學
            "facilitating further education": "可銜接未來升學",
            "further education": "可銜接未來升學",

            # 同儕
            "peer group": "受同儕影響",

            # 師長（保險）
            "teachers": "受師長影響",
            "faculty": "受師長影響",

        },
        "exact": {
            "興趣/適性": "興趣/適性",
            "未來出路/就業": "未來出路/就業",
            "科系聲望/評價": "科系聲望/評價",
            "課程內容/特色": "課程內容/特色",
            "師資": "師資",
            "朋友、同學、學長姐的影響或建議": "朋友、同學、學長姐的影響或建議",
            "考試成績決定": "考試成績決定",
            "地緣關係": "地緣關係",
            "學費/經濟考量": "學費/經濟考量",
            "以上皆非": "以上皆非",
            "其他": "其他",
        },
        "unknown_to_other": True,
    },
    "13本校在我心目中的志願排序": {
        "type": "single",
        "contains": {
            "1st - 3rd": "前3志願",
            "4th - 6th": "第4～6志願",
            "7th - 10th": "第7～10志願",
            "11th": "第11志願以外",
        },
        "exact": {
            "前3志願": "前3志願",
            "第4～6志願": "第4～6志願",
            "第7～10志願": "第7～10志願",
            "第11志願以外": "第11志願以外",
        },
        "unknown_to_other": False,
    },
    "14目前就讀科系在我心目中的志願排序": {
        "type": "single",
        "contains": {
            "1st - 3rd": "前3志願",
            "4th - 6th": "第4～6志願",
            "7th - 10th": "第7～10志願",
            "11th": "第11志願以外",
        },
        "exact": {
            "前3志願": "前3志願",
            "第4～6志願": "第4～6志願",
            "第7～10志願": "第7～10志願",
            "第11志願以外": "第11志願以外",
        },
        "unknown_to_other": False,
    },
    "16目前就讀科系之專業領域":{
        "type": "single",
        "custom": "q16_academic_group",   # 用 custom 指示走函式
    },
    "17目前就讀學制": {
        "type": "single",
        "contains": {
            "daytime": "四年制學士班（日間部）",
            "日間部": "四年制學士班（日間部）",
            "日四技": "四年制學士班（日間部）",

            "進修": "四年制學士班（進修部）",
            "evening": "四年制學士班（進修部）",
            "進二技": "四年制學士班（進修部）",

            "經管進修學士班": "其他",
            "其他": "其他",
            "不知道": "其他",
        },
        "exact": {
            "四年制學士班（日間部）": "四年制學士班（日間部）",
            "四年制學士班（進修部）": "四年制學士班（進修部）",
            "其他": "其他",
        },
        "unknown_to_other": False,
    },
    "18目前就讀科系是否舉辦下列新生活動（可複選）":{
        "type": "multiselect",
        "contains":{
            "orientation held by school": "全校性新生說明會",
            "orientation held by institute": "系新生說明會",
            "orientation held by department": "系新生說明會",
            "orientation held by institute/department": "系新生說明會",
            "welcome activities": "迎新活動",
            "seniors": "直系學長姐自辦活動/聚會",
            "gatherings with seniors": "直系學長姐自辦活動/聚會",
            "其他": "其他",
            "other": "其他",
        },
        "exact": {
            "無":"無",
            "none": "無",
            "no": "無",
        },
        "unknown_to_other": True,
    },
    "30入學至今，我感到比較困擾的問題（可複選）": {
        "type": "multiselect",

        # ✅ contains：用「片段命中」把中英/同義寫法併回你要的類別
        "contains": {
            # ---- 類別A：取得課業學習上的協助（範例：可自行改名） ----
            "資源": "取得課業學習上的協助",
            "輔導": "取得課業學習上的協助",
            "諮詢": "取得課業學習上的協助",
            "tutor": "取得課業學習上的協助",
            "counsel": "取得課業學習上的協助",
            "support": "取得課業學習上的協助",

            # ---- 類別B：課程/教學 ----
            "課業": "課程/教學",
            "課程": "課程/教學",
            "教學": "課程/教學",
            "授課": "課程/教學",
            "curriculum": "課程/教學",
            "course": "課程/教學",
            "teaching": "課程/教學",
            

            # ---- 類別C：人際/同儕 ----
            "同學": "人際/同儕",
            "同儕": "人際/同儕",
            "朋友": "人際/同儕",
            "classmate": "人際/同儕",
            "sex": "人際/同儕",
            "peer": "人際/同儕",
            "friend": "人際/同儕",
            "club": "參與社團活動",
            "teacher": "與老師的互動",
            "family": "與家人的互動",


            # ---- 類別D：生涯/就業/升學 ----
            "升學": "對未來進路的迷惘",
            "就業": "對未來進路的迷惘",
            "職涯": "對未來進路的迷惘",
            "career": "對未來進路的迷惘",
            "employ": "對未來進路的迷惘",
            "further education": "對未來進路的迷惘",
            "future direction": "對未來進路的迷惘",
            

            # ---- 類別E：身心健康 ----
            "壓力": "身心健康",
            "心理": "身心健康",
            "情緒": "身心健康",
            "health": "身心健康",
            "mental": "身心健康",
            "stress": "身心健康",
            "emotion": "身心健康",
            "game": "沉迷網路/網路遊戲等",
            "time": "時間管理",
            "financial": "經濟壓力",

            # ---- 其他/以上皆非（保底）----
            "以上皆非": "以上皆非",
            "none of the above": "以上皆非",
            "其他": "其他",
            "other": "其他",
            "others": "其他",
        },

        # ✅ exact：讓已經是「目標類別名」的值也能穩定保持
        "exact": {
            "取得課業學習上的協助": "取得課業學習上的協助",
            "課程/教學": "課程/教學",
            "人際/同儕": "人際/同儕",
            "對未來進路的迷惘": "對未來進路的迷惘",
            "pressure from academic competition": "身心健康",
            "preparing for class": "課程/教學",
            "身心健康": "身心健康",
            "以上皆非": "以上皆非",
            "其他": "其他",
        },

        # 少數漏網（自由填答）先歸其他，之後再用 UNMAPPED 回補
        "unknown_to_other": True,
    },


    # =========================
    # Q31（多選）
    # =========================
    "31未來生涯規劃最主要的目標為哪一項": {
        "type": "multiselect",
        "contains": {
            "abroad": "國外進修",
            "work": "進入職場",
            "continuing your education": "國內繼續升學",
        },
        "exact": {
            "under planning": "尚無明確規劃",
            "以上皆非": "以上皆非",
            "其他": "其他",
        },
        "unknown_to_other": True,
    },


    # =========================
    # Q32（多選）
    # =========================
    "32就學期間，我「預計的學習重點」（可複選）": {
    "type": "multiselect",

    "contains": {
        # 1 專業能力
        "professional skills": "與就業相關的專業能力",
        "professional skills at work": "與就業相關的專業能力",

        # 2 領導 / 管理
        "leadership": "領導、管理及規劃能力",
        "management": "領導、管理及規劃能力",
        "planning": "領導、管理及規劃能力",

        # 3 電腦文書
        "software used for word processing": "電腦文書處理能力",
        "word processing": "電腦文書處理能力",

        # 4 資訊科技
        "information technology": "資訊科技應用能力",
        "IT application": "資訊科技應用能力",

        # 5 外語
        "foreign language": "外語能力",
        "foreign languages": "外語能力",

        # 6 國際移動
        "transferring internationally": "國際移動力",
        "international mobility": "國際移動力",

        # 7 寫作
        "written communication": "文字撰寫能力",
        "writing": "文字撰寫能力",

        # 8 閱讀
        "reading comprehension": "閱讀能力",

        # 9 表達
        "presentation": "清晰有條理的表達能力",
        "clear and organized presentation": "清晰有條理的表達能力",

        # 10 人際
        "interpersonal communication": "人際溝通能力",

        # 11 團隊 / 合作
        "teamwork": "國際合作能力",
        "collaboration": "國際合作能力",

        # 12 創新
        "innovative": "創新創意思維能力",
        "creative thinking": "創新創意思維能力",

        # 13 創業
        "entrepreneurship": "創業精神與態度",

        # 14 行政
        "administrative coordination": "行政協調能力",

        # 15 解決問題
        "problem detecting": "發現及解決問題的能力",
        "problem solving": "發現及解決問題的能力",

        # 16 執行
        "project implementation": "專業執行能力",

        # 17 分析
        "numerical reasoning": "分析與批判思考能力",
        "logical reasoning": "分析與批判思考能力",
        "critical thinking": "分析與批判思考能力",

        # 18 跨域
        "interdisciplinary": "跨域整合能力",
        "integration": "跨域整合能力",

        # 19 資訊蒐整
        "information collection": "資訊蒐集與整理能力",
        "information processing": "資訊蒐集與整理能力",

        # 20 情緒
        "emotional management": "情緒管理能力",

        # 21 時間
        "time management": "時間管理能力",

        # 22 自主
        "independent learning": "自主學習能力",
        "self learning": "自主學習能力",

        # 23 社會參與
        "social engagement": "社會參與（實踐）能力",
        "practice": "社會參與（實踐）能力",

        # 24 其他
        "other": "其他",
        "其他": "其他",
    },

    "exact": {
        "與就業相關的專業能力": "與就業相關的專業能力",
        "領導、管理及規劃能力": "領導、管理及規劃能力",
        "電腦文書處理能力": "電腦文書處理能力",
        "資訊科技應用能力": "資訊科技應用能力",
        "外語能力": "外語能力",
        "國際移動力": "國際移動力",
        "文字撰寫能力": "文字撰寫能力",
        "閱讀能力": "閱讀能力",
        "清晰有條理的表達能力": "清晰有條理的表達能力",
        "人際溝通能力": "人際溝通能力",
        "國際合作能力": "國際合作能力",
        "創新創意思維能力": "創新創意思維能力",
        "創業精神與態度": "創業精神與態度",
        "行政協調能力": "行政協調能力",
        "發現及解決問題的能力": "發現及解決問題的能力",
        "專業執行能力": "專業執行能力",
        "分析與批判思考能力": "分析與批判思考能力",
        "跨域整合能力": "跨域整合能力",
        "資訊蒐集與整理能力": "資訊蒐集與整理能力",
        "情緒管理能力": "情緒管理能力",
        "時間管理能力": "時間管理能力",
        "自主學習能力": "自主學習能力",
        "社會參與（實踐）能力": "社會參與（實踐）能力",

    },

    "unknown_to_other": True,
}




}

def normalize_column(series, mapping):
    def _norm(x):
        if pd.isna(x):
            return pd.NA
        key = str(x).strip().lower()
        return mapping.get(key, x)  # 不在 map 的，原樣保留
    return series.apply(_norm)

_NUM_RE = re.compile(r"^\s*(\d+)\s*(.*)$")

UNKNOWN_KEYWORDS = [
    "unknown",
    "dont know",
    "don't know",
    "不知",
    "不知道",
    "6 unknown"
]

def is_unknown(s):
    low = s.lower()
    return any(k in low for k in UNKNOWN_KEYWORDS)

def parse_likert_1to5_unknown6(x, unknown_code=6, unknown_label="不知道"):
    if pd.isna(x):
        return (pd.NA, pd.NA)

    s = str(x).strip()
    if not s:
        return (pd.NA, pd.NA)

    #  (1) 最優先：先抓所有 unknown（無論有沒有數字）
    if is_unknown(s):
        return (pd.NA, unknown_label)

    m = _NUM_RE.match(s)

    # (2) 沒有數字開頭
    if not m:
        return (pd.NA, s)

    code = int(m.group(1))

    # (3) 數字是 unknown code（例如 6）
    if code == unknown_code:
        return (pd.NA, unknown_label)

    if 1 <= code <= 5:
        return (code, str(code))

    return (pd.NA, str(code))
