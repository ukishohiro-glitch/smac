# -*- coding: utf-8 -*-
# main.py — ①大分類ラジオ ②片引き/引分けラジオ ③部材品名ラジオ ④定型文チェック ⑤Excel出力(お見積書(明細)) 対応版
from __future__ import annotations
import os, os.path as osp, secrets, math, re, unicodedata
from datetime import datetime, date
import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook

# ===== ユーティリティ =====
def normalize_string(s):
    if isinstance(s, str):
        s = s.replace("\u3000", " ")
        s = unicodedata.normalize("NFKC", s)
        s = re.sub(r"\s+", " ", s.strip())
    return s

def norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty: return pd.DataFrame()
    df = df.copy()
    df.columns = [normalize_string(c) for c in df.columns]
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(normalize_string)
    return df

NUM_RE = re.compile(r"[-+]?\d[\d,]*\.?\d*")
def parse_float(x):
    if x is None: return None
    m = NUM_RE.search(str(x))
    if not m: return None
    try: return float(m.group(0).replace(",", ""))
    except: return None

def parse_length_mm(x):
    if x is None: return None
    s = normalize_string(str(x)).lower()
    m_m = re.search(r"(\d[\d,]*\.?\d*)\s*m(?!m)", s)
    if m_m:
        return int(math.ceil(float(m_m.group(1).replace(",", ""))*1000))
    v = parse_float(s)
    if v is None: return None
    if "mm" in s: return int(round(v))
    return int(round(v if v>=1000 else v*1000))

def ceil100(x: float) -> int: return int(math.ceil(x/100.0)*100)

def pick_col(df: pd.DataFrame, candidates) -> str | None:
    for c in candidates:
        if c in df.columns: return c
    norm = {normalize_string(c): c for c in df.columns}
    for c in candidates:
        if c in norm: return norm[c]
    return None

def read_xlsx(path: str, sheet: str | None = None) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=sheet) if sheet else pd.read_excel(path)
        return norm_cols(df)
    except Exception:
        return pd.DataFrame()

# ===== マスタ読込 =====
def load_master():
    xls = "master.xlsx"
    def R(sn, add_cols=None):
        try:
            df = pd.read_excel(xls, sheet_name=sn); df = norm_cols(df)
            if add_cols:
                for c in add_cols:
                    if c not in df.columns: df[c] = ""
            return df
        except Exception: return pd.DataFrame()

    df_clients  = R("得意先一覧")
    df_curtain  = R("カーテン", ["大分類","中分類","小分類"])
    df_perf     = R("カーテン性能", ["中分類","性能"])
    df_ma       = R("MA型単価表")
    df_mb_tbl   = R("MB型単価表")
    df_mc       = R("MC型単価表")
    df_me_curt  = read_xlsx(xls, "ME型単価表")
    df_me_motor = read_xlsx(xls, "ME型OP")
    df_op       = R("OP", ["OP名称","金額","方向"])
    df_gap      = R("隙間シート")
    parts_sheets = ["カーテンレール","取手付間仕切ポール","中間ポール","アルミ押えバー","間仕切ネットBOXバー","落し","その他"]
    df_parts = {sn: R(sn) for sn in parts_sheets}
    df_gen = read_xlsx("原反価格表.xlsx")
    if df_gen.empty: df_gen = R("原反価格")
    return (df_clients, df_curtain, df_perf,
            df_ma, df_mb_tbl, df_mc, df_me_curt, df_me_motor,
            df_op, df_gap, df_parts, df_gen)

# ===== S・MAC計算（簡略） =====
def extract_thickness(text: str) -> float | None:
    m = re.search(r"(\d+(?:\.\d+)?)\s*t", str(text or "")); 
    if m:
        try: return float(m.group(1))
        except: return None
    return None
def smac_estimate(middle_name: str, open_method: str, W: int, H: int, cnt: int,
                  df_gen: pd.DataFrame, df_op: pd.DataFrame, picked_ops: list[dict]):
    """戻り値: dict(ok, sell_one, sell_total, note_ops, breakdown)
       ※OP金額はカーテンに内包（サマリ行は名称のみ表示）"""
    res = {"ok": False, "msg": "", "sell_one": 0, "sell_total": 0, "note_ops": [], "breakdown": {}}
    if not middle_name or W<=0 or H<=0 or cnt<=0 or df_gen.empty:
        res["msg"] = "S・MAC：中分類/寸法/数量/原反価格を確認してください。"
        return res

    name_col = pick_col(df_gen, ["製品名","品名","名称"]) or df_gen.columns[0]
    w_col    = pick_col(df_gen, ["幅","巾","原反幅(mm)","原反幅"]) or df_gen.columns[1]
    u_col    = pick_col(df_gen, ["単価","上代","価格","金額"]) or df_gen.columns[2]
    t_col    = pick_col(df_gen, ["厚み","厚さ","t"])

    hit = df_gen[df_gen[name_col]==middle_name]
    if hit.empty:
        hit = df_gen[df_gen[name_col].astype(str).str.contains(re.escape(middle_name), na=False)]
    if hit.empty:
        res["msg"] = f"S・MAC：原反価格に『{middle_name}』が見つかりません。"
        return res

    gen_width = parse_float(hit.iloc[0][w_col])
    gen_price = parse_float(hit.iloc[0][u_col])
    thick = extract_thickness(hit.iloc[0][t_col]) if t_col else extract_thickness(hit.iloc[0][name_col])
    if not gen_width or not gen_price:
        res["msg"] = "S・MAC：原反幅または単価が不正です。"
        return res

    # 寸法・枚数
    if open_method == "片引き":
        cur_w = W * 1.05; panels = 1
    else:
        cur_w = (W/2) * 1.05; panels = 2
    cur_h = H + 50

    # 原価計算（1間口あたり）
    length_per_panel_m = (cur_h * 1.2) / 1000.0       # 1パネルの必要長(m)
    joints = math.ceil(cur_w / gen_width)            # 巾継ぎ本数
    raw_one = gen_price * length_per_panel_m * joints * panels       # 原反
    cutting_one = (2000 if joints <= 3 else 3000) * panels           # 裁断
    hem_unit = 450 if (thick is not None and thick <= 0.3) else 550  # 三巻など
    hem_perimeter_m = (cur_w + cur_w + cur_h + cur_h) / 1000.0
    hem_total = math.ceil(hem_perimeter_m) * hem_unit * panels * cnt # ←合計
    hem_one = hem_total / cnt                                         # 1間口あたり

    # OP（1000mm切上げ・合計はカーテンに内包）
    note_ops = []; op_total = 0
    for op in picked_ops or []:
        name = normalize_string(op.get("OP名称",""))
        unit = int(parse_float(op.get("金額")) or 0)
        dire = normalize_string(op.get("方向","")).upper()
        if not name or unit <= 0:
            continue
        base_mm = cur_w if dire in ["W","横","X"] else cur_h
        units_1000 = math.ceil(base_mm/1000.0)
        sub = units_1000 * unit * panels * cnt
        op_total += sub
        note_ops.append(name)
    op_one = op_total / cnt if cnt else 0

    # 原価合計→売値（既存ロジックに合わせ粗利40%: /0.6）
    genka_total = raw_one*cnt + cutting_one*cnt + hem_total + op_total
    sell_one = ceil100((raw_one + cutting_one + hem_one + op_one) / 0.6)
    sell_total = ceil100(genka_total / 0.6)

    breakdown = {
        "原反(材料)": int(round(raw_one)),
        "裁断":       int(round(cutting_one)),
        "周囲三巻/縫製": int(round(hem_one)),
        "OP合計":    int(round(op_one)),
    }
    res.update({"ok": True, "sell_one": int(sell_one), "sell_total": int(sell_total),
                "note_ops": note_ops, "breakdown": breakdown})
    return res


# ===== エア・セーブ簡易 =====
def pick_price_col(df: pd.DataFrame):
    for c in ["㎡単価","平米単価","m2単価","単価","上代","価格","金額"]:
        if c in df.columns: return c
    return None

def area_price(df_area: pd.DataFrame, item_name: str, perf: str, W: int, H: int, CNT: int):
    res = {"ok": False, "msg": "", "unit":0, "price_one":0, "total":0}
    if df_area.empty or not item_name or W<=0 or H<=0 or CNT<=0:
        res["msg"] = "単価表・品名・寸法・数量を確認してください。"; return res
    name_col = pick_col(df_area, ["品名","製品名","名称","品番","型式"]) or df_area.columns[0]
    perf_col = pick_col(df_area, ["性能","特性"])
    unit_col = pick_price_col(df_area)
    if unit_col is None: res["msg"] = "単価列が見つかりません。"; return res

    rows = df_area[df_area[name_col]==item_name]
    if perf_col and perf: rows = rows[rows[perf_col]==perf]
    if rows.empty: res["msg"] = "対象行が見つかりません。"; return res

    unit = parse_float(rows.iloc[0][unit_col])
    if not unit: res["msg"] = "単価が数値化できません。"; return res

    area = (W/1000)*(H/1000)
    price_one = ceil100(area*unit); total = int(price_one*CNT)
    res.update({"ok":True, "unit": int(round(unit)), "price_one": int(price_one), "total": int(total)})
    return res

def mc_slide_rail_price(W: int, CNT: int) -> int:
    rail_len_mm = int(math.ceil((W*2)/2000.0)*2000)
    return int((rail_len_mm/1000.0)*7400)*CNT

def fixed_price(df_fixed: pd.DataFrame, item_name: str) -> int:
    if df_fixed.empty or not item_name: return 0
    name_col = pick_col(df_fixed, ["品名","製品名","名称"]) or df_fixed.columns[0]
    price_col = pick_price_col(df_fixed) or pick_col(df_fixed, ["固定価格","単価","価格","金額"])
    if price_col is None: return 0
    row = df_fixed[df_fixed[name_col]==item_name]
    if row.empty: return 0
    v = parse_float(row.iloc[0][price_col])
    return int(ceil100(v)) if v else 0

# ===== 部材ヘルパ =====
PART_LENGTH_RULE = {
    "カーテンレール": "width",
    "取手付間仕切ポール": "height",
    "中間ポール": "height",
    "アルミ押えバー": "height",
    "間仕切ネットBOXバー": "height",
}
REQ_COLS_MAP = {
    "品名": ["品名","製品名","名称","品番"],
    "m単価": ["m単価","/m","1m単価","１m単価","m当たり","mあたり"],
    "セット長(mm)": ["セット長(mm)","セット長","定尺","規格長","サイズ","寸法","長さ","長さ(mm)"],
    "セット価格": ["セット価格","セット金額"],
    "固定価格": ["固定価格","単価","価格","金額","上代"],
}
def get_length_columns(df: pd.DataFrame) -> list[int]:
    lens = []
    for c in df.columns:
        try:
            v = int(str(c).replace(",", ""))
            if v>=1000: lens.append(v)
        except: pass
    return sorted(set(lens))

def pick_len_col(length_cols: list[int], need_mm: int) -> int | None:
    if not length_cols: return None
    for L in length_cols:
        if L>=need_mm: return L
    return length_cols[-1]

def map_required_cols(df: pd.DataFrame) -> dict:
    mp = {}
    for formal, cands in REQ_COLS_MAP.items():
        col = pick_col(df, cands)
        if col: mp[formal] = col
    return mp

def price_part_row(part_name: str, row: pd.Series, W: int, H: int, df_sheet: pd.DataFrame) -> tuple[int, str, float]:
    length_cols = get_length_columns(df_sheet)
    if length_cols:
        basis = 0
        if PART_LENGTH_RULE.get(part_name)=="width" and W>0:  basis=int(W)
        elif PART_LENGTH_RULE.get(part_name)=="height" and H>0: basis=int(H)
        if basis>0:
            col_mm = pick_len_col(length_cols, basis)
            if col_mm is not None:
                price_cell = row.get(col_mm) if col_mm in row else row.get(str(col_mm))
                price_raw = parse_float(price_cell)
                if price_raw and price_raw>0:
                    unit = ceil100(price_raw*0.5)
                    label = f"{int(col_mm/1000)}m相当"
                    return int(unit), label, float(col_mm)/1000.0
    mp = map_required_cols(pd.DataFrame([row]))
    if "m単価" in mp:
        per_m = parse_float(row[mp["m単価"]])
        if per_m and per_m>0:
            if PART_LENGTH_RULE.get(part_name)=="width" and W>0:
                mcount = math.ceil(W/1000.0)
            elif PART_LENGTH_RULE.get(part_name)=="height" and H>0:
                mcount = math.ceil((H/1000.0)*2)/2
            else: mcount = 0
            unit = int(ceil100(per_m*0.5) * (mcount if mcount else 0))
            label = f"{mcount:g}m分" if mcount else ""
            return unit, label, float(mcount or 0)
    if "セット長(mm)" in mp and "セット価格" in mp:
        L = parse_length_mm(row[mp["セット長(mm)"]]); P = parse_float(row[mp["セット価格"]])
        if L and L>=1000 and P and P>0:
            unit = ceil100(P*0.5); label=f"{int(L/1000)}mセット"
            return int(unit), label, float(L)/1000.0
    if "固定価格" in mp:
        P = parse_float(row[mp["固定価格"]])
        if P and P>0: return int(ceil100(P*0.5)), "", 0.0
    return 0, "", 0.0

# ===== 見積番号 =====
FY_BASE_DATE = date(2024, 10, 1)  # 37期起点
FY_BASE_TERM = 37
def get_term_for(today: date) -> int:
    term = FY_BASE_TERM + (today.year - FY_BASE_DATE.year)
    if today < date(today.year, 10, 1): term -= 1
    return term
def generate_estimate_no(user_code: str, today: date, seen_serials: set[str]) -> str:
    uc = normalize_string(user_code or "").upper() or "UC"
    term = get_term_for(today); mm = f"{today.month:02d}"
    while True:
        sfx = f"{secrets.randbelow(10000):04d}"
        candidate = f"{uc}{term}{mm}-{sfx}"
        if candidate not in seen_serials:
            seen_serials.add(candidate); return candidate

# ===== 画面セットアップ =====
st.set_page_config(layout="wide", page_title="お見積書作成システム")

# CSS少し詰め
st.markdown("""
<style>
section.main > div.block-container { padding-top: 8px; padding-bottom: 10px; padding-left: 10px; padding-right: 10px; }
div[data-testid="stHorizontalBlock"] { gap: 8px !important; }
div[data-testid="column"] { padding-left: 6px; padding-right: 6px; }
.sec-h { font-size:16pt; font-weight:400; margin: 6px 0 4px; }
.row-compact { display: grid; grid-template-columns: 1.4fr 0.45fr 0.7fr 0.8fr 0.28fr; column-gap: 8px; align-items: end; }
.sticky-wrap { position: sticky; top: 0; z-index: 999; background: var(--background-color); padding: 6px 6px 8px; border-bottom: 1px solid rgba(128,128,128,.35); }
.hr-thin { border-top: 1px solid rgba(128,128,128,.35); margin: 6px 0 10px; }
</style>
""", unsafe_allow_html=True)
def sec_title(text: str): st.markdown(f'<div class="sec-h">{text}</div>', unsafe_allow_html=True)

# マスタ
(df_clients, df_curtain, df_perf,
 df_ma, df_mb_tbl, df_mc, df_me_curt, df_me_motor,
 df_op, df_gap, df_parts, df_gen) = load_master()

# セッション
today = date.today()
if "seen_serials" not in st.session_state: st.session_state.seen_serials = set()
if "user_code" not in st.session_state:   st.session_state.user_code = "UC"
if "estimate_no" not in st.session_state: st.session_state.estimate_no = generate_estimate_no(st.session_state.user_code, today, st.session_state.seen_serials)
if "openings" not in st.session_state:    st.session_state.openings = [{"id":1}]
if "file_title_manual" not in st.session_state: st.session_state.file_title_manual = False
if "file_title" not in st.session_state:        st.session_state.file_title = datetime.today().strftime("%m%d")

# 定型文（④）
PHRASES = [
    "金具：スチール", "金具：ステンレス",
    "戸先側：取手付間仕切りポール", "戸尻側：フラットバー固定", "戸尻側：取手付間仕切りポール", "戸尻側：吊下げ固定",
    "※ シートは収縮を考慮し長めでの出荷となります。現場で裾カット調整してください。",
    "※ カーテン下端は長めでの出荷となります。現場で裾カット調整してください。",
    "※ 下地別途。取付用のビス等は別途。", 
    "※ カーテンと間仕切りポール・中間ポールは組込済みです。取手・落し・マグネットは現場で取り付けてください。",
]

# ===== ヘッダ =====
st.markdown('<div class="sticky-wrap">', unsafe_allow_html=True)
sec_title("お見積書作成システム")
c1, c2, c3 = st.columns([0.9, 1.2, 0.6])
with c1: st.text_input("担当者コード", value=st.session_state.user_code, key="user_code")
with c2: st.text_input("見積番号", value=st.session_state.estimate_no, key="estimate_no")
with c3:
    if st.button("見積番号を再生成", key="regen_no"):
        st.session_state.estimate_no = generate_estimate_no(st.session_state.user_code, today, st.session_state.seen_serials)
        st.rerun()

d1,d2,d3,d4 = st.columns([1.1,1.1,1.1,1.1])
clients = df_clients["社名"].dropna().unique().tolist() if "社名" in df_clients.columns else []
with d1:
    sel_client = st.selectbox("得意先名", [""]+clients, key="client")
with d2:
    branches = df_clients[df_clients["社名"]==sel_client]["支店名"].dropna().unique().tolist() if sel_client and "支店名" in df_clients.columns else []
    sel_branch = st.selectbox("支店名", [""]+branches, key="branch")
with d3:
    offices = df_clients[(df_clients["社名"]==sel_client)&(df_clients["支店名"]==sel_branch)]["営業所名"].dropna().unique().tolist() if sel_client and sel_branch and "営業所名" in df_clients.columns else []
    sel_office = st.selectbox("営業所名", [""]+offices, key="office")
with d4:
    contact = st.text_input("担当者名", key="pic")
e1,e2 = st.columns([2.6,0.6])
with e1:
    pj_name = st.text_input("物件名", key="pj")
    if not st.session_state.file_title_manual:
        st.session_state.file_title = f"{datetime.today().strftime('%m%d')}{pj_name}" if pj_name else datetime.today().strftime('%m%d')
with e2:
    st.text_input("作成日", value=today.strftime("%Y/%m/%d"), key="created_disp", disabled=True)
st.markdown('<div class="hr-thin"></div>', unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

# 合計サマリ
overall_total = 0
overall_items = []
def overall_total_update(v): 
    global overall_total
    overall_total += int(v or 0)

# ===== 間口 =====
def render_opening(idx: int):
    pref = f"o{idx}_"
    # 行A
    a1,a2,a3,a4 = st.columns([0.7,1.0,1.0,0.7])
    with a1: mark = st.text_input("符号", key=pref+"mark")
    with a2: W = st.number_input("間口W (mm)", min_value=0, value=0, step=50, key=pref+"w")
    with a3: H = st.number_input("間口H (mm)", min_value=0, value=0, step=50, key=pref+"h")
    with a4: CNT = st.number_input("数量", min_value=1, value=1, step=1, key=pref+"cnt")

    sec_title("カーテン入力")
    # 行B：①大分類ラジオ
    b1,b2,b3 = st.columns([1.0,1.0,1.0])
    with b1:
        large_list = []
        if "大分類" in df_curtain.columns:
            large_list = [x for x in df_curtain["大分類"].dropna().unique().tolist() if x]
        if "エア・セーブ" not in large_list:
            large_list.append("エア・セーブ")
        large = st.radio("カーテン大分類", [""]+large_list, key=pref+"large", horizontal=True)
    air_type = None; middle = small = perf = ""; rib_note = ""

    # B2/B3
    with b2:
        if large and large != "エア・セーブ":
            mids = df_curtain[df_curtain["大分類"]==large]["中分類"].dropna().unique().tolist() if "中分類" in df_curtain.columns else []
            middle = st.selectbox("カーテン中分類", [""]+mids, key=pref+"mid")
    with b3:
        if large == "エア・セーブ":
            air_label = st.radio("型式（MA・MB・MC・ME）", ["","MA型折りたたみ式","MB型固定式","MC型スライド式","ME型電動式"], key=pref+"airtype", horizontal=True)
            air_type = air_label[:2] if air_label else None
        else:
            if middle:
                small_opts = df_curtain[(df_curtain["大分類"]==large)&(df_curtain["中分類"]==middle)]["小分類"].dropna().unique().tolist() if "小分類" in df_curtain.columns else []
                small = st.selectbox("カーテン小分類", [""]+small_opts, key=pref+"small")

    # 行C：②片引き/引分けラジオ + 付帯
    c1,c2,c3,c4 = st.columns([0.8,0.8,1.2,1.0])
    with c1:
        if large in ["S・MACカーテン"] or (large=="エア・セーブ" and air_type in ["MC","ME"]):
            open_mtd = st.radio("片引き/引分け", ["","片引き","引分け"], key=pref+"open", horizontal=True)
        else:
            open_mtd = ""
    with c2:
        if large=="エア・セーブ" and air_type in ["MB","MC","ME"]:
            rib_note = st.radio("リブ付き/リブ無し", ["","リブ付き","リブ無し"], key=pref+"rib", horizontal=True)
    with c3:
        air_item = ""
        if large=="エア・セーブ" and air_type:
            if air_type=="MA":
                name_col = pick_col(df_ma, ["品名","製品名","名称","品番","型式"]) or (df_ma.columns[0] if not df_ma.empty else None)
                items = df_ma[name_col].dropna().unique().tolist() if name_col else []
                air_item = st.selectbox("エア・セーブ品名", [""]+items, key=pref+"ma_item")
            elif air_type=="MB":
                name_col = pick_col(df_mb_tbl, ["品名","製品名","名称","品番","型式"]) or (df_mb_tbl.columns[0] if not df_mb_tbl.empty else None)
                items = df_mb_tbl[name_col].dropna().unique().tolist() if name_col else []
                air_item = st.selectbox("エア・セーブ品名", [""]+items, key=pref+"mb_item")
            elif air_type=="MC":
                name_col = pick_col(df_mc, ["品名","製品名","名称","品番","型式"]) or (df_mc.columns[0] if not df_mc.empty else None)
                items = df_mc[name_col].dropna().unique().tolist() if name_col else []
                air_item = st.selectbox("エア・セーブ品名", [""]+items, key=pref+"mc_item")
            elif air_type=="ME":
                curt_df = df_me_curt.copy() if not df_me_curt.empty else df_mc.copy()
                curt_name_col = pick_col(curt_df, ["品名","製品名","名称","品番","型式"]) or (curt_df.columns[0] if not curt_df.empty else None)
                items = curt_df[curt_name_col].dropna().unique().tolist() if curt_name_col else []
                air_item = st.selectbox("エア・セーブ品名（カーテン）", [""]+items, key=pref+"me_curt")
    with c4:
        if large=="エア・セーブ":
            perf_all = df_perf["性能"].dropna().unique().tolist() if "性能" in df_perf.columns else []
            perf = st.selectbox("カーテン性能", [""]+perf_all, key=pref+"perf")
        elif large and large!="S・MACカーテン":
            perf_opts = df_perf[df_perf["中分類"]==middle]["性能"].dropna().unique().tolist() if "中分類" in df_perf.columns else []
            perf = st.selectbox("カーテン性能", [""]+perf_opts, key=pref+"perf2")

    # S・MAC OP 簡易選択（任意）
    picked_ops = []
    if large=="S・MACカーテン" and not df_op.empty and all(c in df_op.columns for c in ["OP名称","金額","方向"]):
        st.caption("S・MAC OP（任意／金額はカーテンに内包）")
        names = df_op["OP名称"].dropna().unique().tolist()
        cols = st.columns(3)
        for i, nm in enumerate(names):
            with cols[i%3]:
                if st.checkbox(nm, key=pref+f"smac_op_{i}"):
                    picked_ops.append(df_op[df_op["OP名称"]==nm].iloc[0].to_dict())

    # 計算→サマリ
    if W>0 and H>0 and CNT>0:
        if large=="S・MACカーテン":
            sm = smac_estimate(middle or "", st.session_state.get(pref+"open") or "片引き",
                               W, H, CNT, df_gen, df_op, picked_ops)
            if sm["ok"]:
                note = f"W{W}×H{H}mm"
                if sm["note_ops"]: note += "／OP：" + "・".join(sm["note_ops"])
                overall_items.append({
# （S・MACサマリ追加の直後）
if sm.get("breakdown"):
    with st.expander("原価構成（1間口あたり）", expanded=False):
        b = sm["breakdown"]
        dfb = pd.DataFrame(
            {"項目": list(b.keys()), "金額": [int(b[k]) for k in b.keys()]}
        )
        st.dataframe(dfb, use_container_width=True, hide_index=True)
        st.caption("※ OP金額はカーテン金額に内包・サマリには名称のみ表示")

                    "品名": "S・MACカーテン" + (f" {middle}" if middle else "") + (f" {st.session_state.get(pref+'open')}" if st.session_state.get(pref+'open') else ""),
                    "数量": CNT, "単位":"式", "単価": sm["sell_one"], "小計": sm["sell_total"],
                    "種別":"S・MAC", "備考": (f"符号:{mark}／" if mark else "") + note
                }); overall_total_update(sm["sell_total"])
        if large=="エア・セーブ" and air_type:
            if air_type=="MA" and air_item and perf:
                r = area_price(df_ma, air_item, perf, W, H, CNT)
                if r["ok"]:
                    overall_items.append({
                        "品名": f"エア・セーブ MA型折りたたみ式 {air_item}",
                        "数量": CNT, "単位":"式", "単価": r["price_one"], "小計": r["total"],
                        "種別":"エア・セーブMA", "備考": (f"符号:{mark}／" if mark else "") + f"W{W}×H{H}mm"
                    }); overall_total_update(r["total"])
            if air_type=="MB" and air_item and perf:
                r = area_price(df_mb_tbl, air_item, perf, W, H, CNT)
                if r["ok"]:
                    overall_items.append({
                        "品名": f"エア・セーブ MB型固定式 {air_item}",
                        "数量": CNT, "単位":"式", "単価": r["price_one"], "小計": r["total"],
                        "種別":"エア・セーブMB",
                        "備考": (f"符号:{mark}／" if mark else "") + f"W{W}×H{H}mm" + (f"／{st.session_state.get(pref+'rib')}" if st.session_state.get(pref+'rib') else "")
                    }); overall_total_update(r["total"])
            if air_type=="MC" and air_item and perf:
                r = area_price(df_mc, air_item, perf, W, H, CNT)
                if r["ok"]:
                    rail = mc_slide_rail_price(W, CNT)
                    total = r["total"] + rail
                    overall_items.append({
                        "品名": f"エア・セーブ MC型スライド式 {air_item}",
                        "数量": CNT, "単位":"式", "単価": r["price_one"], "小計": r["price_one"]*CNT,
                        "種別":"エア・セーブMC",
                        "備考": (f"符号:{mark}／" if mark else "") + f"W{W}×H{H}mm"
                                 + (f"／{st.session_state.get(pref+'rib')}" if st.session_state.get(pref+'rib') else "")
                                 + (f"／{st.session_state.get(pref+'open')}" if st.session_state.get(pref+'open') else "")
                    })
                    overall_items.append({
                        "品名": "スライドレール", "数量": 1, "単位":"式", "単価": rail, "小計": rail,
                        "種別":"エア・セーブMC", "備考": "W×2を2000mm刻み"
                    })
                    overall_total_update(total)
            if air_type=="ME" and perf:
                total_me = 0
                if air_item:
                    curt_df = df_me_curt.copy() if not df_me_curt.empty else df_mc.copy()
                    r = area_price(curt_df, air_item, perf, W, H, CNT)
                    if r["ok"]:
                        overall_items.append({
                            "品名": f"エア・セーブ ME型電動式 カーテン {air_item}",
                            "数量": CNT, "単位":"式", "単価": r["price_one"], "小計": r["total"],
                            "種別":"エア・セーブME(カーテン)",
                            "備考": (f"符号:{mark}／" if mark else "") + f"W{W}×H{H}mm"
                        }); total_me += r["total"]
                if st.session_state.get(pref+"me_motor"):
                    mu = fixed_price(df_me_motor, st.session_state[pref+"me_motor"])
                    if mu>0:
                        overall_items.append({
                            "品名": f"エア・セーブ ME型電動式 駆動部 {st.session_state[pref+'me_motor']}",
                            "数量": CNT, "単位":"式", "単価": mu, "小計": mu*CNT,
                            "種別":"エア・セーブME(駆動部)", "備考": ""
                        }); total_me += mu*CNT
                if total_me>0: overall_total_update(total_me)

    # ===== 部材入力：③品名ラジオ =====
    show_parts = (st.session_state.get(pref+"large")=="S・MACカーテン") or (st.session_state.get(pref+"large")=="エア・セーブ" and (st.session_state.get(pref+"airtype") or "").startswith("MA"))
    if show_parts:
        st.markdown("##### 部材入力")
        for sheet_name, dfp in df_parts.items():
            rows_key = pref + f"{sheet_name}_rows"
            if rows_key not in st.session_state: st.session_state[rows_key] = [{"item":"", "qty":1}]
            # セクション見出し＋追加
            hcol1, hcol2 = st.columns([0.92, 0.08])
            with hcol1: st.caption(f"【{sheet_name}】")
            with hcol2:
                if st.button("＋", key=pref+f"{sheet_name}_add"):
                    st.session_state[rows_key].append({"item":"", "qty":1}); st.rerun()

            name_col = pick_col(dfp, ["品名","製品名","名称","品番"]) or (dfp.columns[0] if not dfp.empty else None)
            names = dfp[name_col].dropna().unique().tolist() if name_col else []
            updated_rows = []
            for j, rowdata in enumerate(st.session_state[rows_key]):
                st.markdown('<div class="row-compact">', unsafe_allow_html=True)
                col1, col2, col3, col4, col5 = st.columns([1.4,0.45,0.7,0.8,0.28])
                with col1:
                    # ③ ラジオボタン化
                    item_opts = [""] + names
                    current = rowdata["item"] if rowdata["item"] in item_opts else ""
                    sel = st.radio(f"品名 {j+1}", item_opts, index=item_opts.index(current), key=pref+f"{sheet_name}_item_{j}_radio")
                with col2:
                    qty = st.number_input("数量", min_value=1, value=int(rowdata["qty"]), step=1, key=pref+f"{sheet_name}_qty_{j}")
                unit, label, used_m = (0,"",0.0)
                if sel:
                    src_row = dfp[dfp[name_col]==sel].iloc[0] if not dfp.empty else None
                    if src_row is not None: unit, label, used_m = price_part_row(sheet_name, src_row, W, H, dfp)
                with col3: st.text_input("単価（自動）", value=str(unit), disabled=True, key=pref+f"{sheet_name}_unit_{j}")
                with col4: st.text_input("使用m数（自動）", value=(f"{used_m:g}" if used_m else ""), disabled=True, key=pref+f"{sheet_name}_m_{j}")
                with col5: delete = st.button("×", key=pref+f"{sheet_name}_del_{j}")
                st.markdown('</div>', unsafe_allow_html=True)

                if not delete:
                    updated_rows.append({"item": sel, "qty": qty})
                    if sel and unit>0:
                        subtotal = unit*qty
                        overall_items.append({
                            "品名": sel, "数量": qty, "単位":"式", "単価": unit, "小計": subtotal,
                            "種別":"部材", "備考": (f"符号:{mark}" if mark else "") + (f"／{label}" if label else "")
                        })
                        overall_total_update(subtotal)
            st.session_state[rows_key] = updated_rows

        # ④ 定型文（部材の下）
        st.markdown("##### 定型文")
        phr_sel = []
        colL, colR = st.columns(2)
        half = (len(PHRASES)+1)//2
        with colL:
            for i, p in enumerate(PHRASES[:half]):
                if st.checkbox(p, key=pref+f"phr_L_{i}"): phr_sel.append(p)
        with colR:
            for i, p in enumerate(PHRASES[half:]):
                if st.checkbox(p, key=pref+f"phr_R_{i}"): phr_sel.append(p)
        if phr_sel:
            overall_items.append({
                "品名": "（定型文）", "数量": "", "単位":"", "単価": "", "小計":"", "種別":"メモ",
                "備考": " / ".join(phr_sel)
            })
    else:
        st.caption("※このカーテン構成では部材の追加は行いません。")

# 描画
for i, op in enumerate(st.session_state.openings, start=1):
    with st.expander(f"間口 {i}", expanded=True):
        render_opening(i)
    cols = st.columns([0.16, 0.84])
    if cols[0].button("この間口を削除", key=f"del_{i}") and len(st.session_state.openings)>1:
        st.session_state.openings.pop(i-1); st.rerun()
if st.button("＋ 間口を追加", key="add_opening"):
    st.session_state.openings.append({"id": len(st.session_state.openings)+1}); st.rerun()

# ===== サマリ =====
st.markdown("---")
sec_title("見積サマリ")
if overall_items:
    df_summary = pd.DataFrame(overall_items, columns=["品名","数量","単位","単価","小計","種別","備考"])
    st.dataframe(df_summary, use_container_width=True, hide_index=True)
else:
    st.info("明細がありません。間口を追加し、カーテン／部材を入力してください。")

sec_title("見積金額")
st.metric("税抜合計", f"¥{overall_total:,}")

# ===== ⑤ Excel出力：「お見積書（明細）」へ =====
def header_dict():
    created = st.session_state.get("created_disp") or datetime.today().strftime("%Y/%m/%d")
    return {
        "estimate_no":   st.session_state.get("estimate_no",""),
        "date":          created.replace("/","-"),
        "customer_name": st.session_state.get("client",""),
        "branch_name":   st.session_state.get("branch",""),
        "office_name":   st.session_state.get("office",""),
        "person_name":   st.session_state.get("pic",""),
        "project_name":  st.session_state.get("pj",""),
    }

def export_to_detail_xlsx(out_path: str, header: dict, items: list[dict], template_path: str | None = None):
    # 既存テンプレートに「お見積書（明細）」があれば置き換え、無ければ新規作成
    ws_name = "お見積書（明細）"
    if template_path and osp.exists(template_path):
        wb = load_workbook(template_path)
        if ws_name in wb.sheetnames:
            wb.remove(wb[ws_name])
        ws = wb.create_sheet(ws_name)
    else:
        wb = Workbook()
        ws = wb.active; ws.title = ws_name

    # ヘッダー情報
    ws["A1"] = "見積番号"; ws["B1"] = header.get("estimate_no","")
    ws["A2"] = "作成日";   ws["B2"] = header.get("date","")
    ws["A3"] = "得意先";   ws["B3"] = header.get("customer_name","")
    ws["A4"] = "支店";     ws["B4"] = header.get("branch_name","")
    ws["A5"] = "営業所";   ws["B5"] = header.get("office_name","")
    ws["A6"] = "担当者";   ws["B6"] = header.get("person_name","")
    ws["A7"] = "物件名";   ws["B7"] = header.get("project_name","")

    start = 9
    cols = ["品名","数量","単位","単価","小計","種別","備考"]
    for j,c in enumerate(cols, start=1): ws.cell(row=start, column=j, value=c)
    r = start + 1
    for it in items:
        ws.cell(r,1,it.get("品名",""))
        ws.cell(r,2,it.get("数量",""))
        ws.cell(r,3,it.get("単位",""))
        ws.cell(r,4,it.get("単価",""))
        ws.cell(r,5,it.get("小計",""))
        ws.cell(r,6,it.get("種別",""))
        ws.cell(r,7,it.get("備考",""))
        r += 1
    wb.save(out_path)

# 保存UI（CSVは従来通り残しつつ、Excel出力を追加）
sec_title("保存")
c1, c2 = st.columns([0.6, 0.4])
save_dir = "./data"; os.makedirs(save_dir, exist_ok=True)
with c1:
    st.checkbox("保存ファイル名を手動で編集する", value=st.session_state.file_title_manual, key="file_title_manual")
    st.text_input("保存ファイル名", value=st.session_state.file_title, key="file_title", disabled=not st.session_state.file_title_manual)
    if st.button("CSV保存", key="save_csv"):
        meta = {
            "見積番号": st.session_state.estimate_no, "作成日": today.strftime("%Y/%m/%d"),
            "ファイル名": st.session_state.file_title, "得意先": st.session_state.get("client",""),
            "支店": st.session_state.get("branch",""), "営業所": st.session_state.get("office",""),
            "担当者": st.session_state.get("pic",""), "物件名": st.session_state.get("pj",""),
        }
        df_meta = pd.DataFrame([meta])
        df_detail = pd.DataFrame(overall_items) if overall_items else pd.DataFrame(columns=["品名","数量","単位","単価","小計","種別","備考"])
        filepath = osp.join(save_dir, f"{st.session_state.file_title}.csv")
        with open(filepath, "w", encoding="utf-8-sig", newline="") as f:
            df_meta.to_csv(f, index=False); f.write("\n"); df_detail.to_csv(f, index=False)
        st.success(f"保存しました: {filepath}")

with c2:
    st.markdown("**Excel保存（お見積書（明細））**")
    if st.button("Excel保存（お見積書（明細））", key="excel_save_btn"):
        if not overall_items:
            st.error("明細がありません。"); 
        else:
            header = header_dict()
            out = osp.join(save_dir, f"{st.session_state.file_title}_お見積書（明細）.xlsx")
            tpl = osp.join(osp.dirname(__file__), "お見積書（明細）.xlsx")  # あれば利用
            try:
                export_to_detail_xlsx(out, header, overall_items, template_path=tpl if osp.exists(tpl) else None)
                st.success(f"Excelを保存しました：{out}")
                with open(out, "rb") as f:
                    st.download_button("ダウンロード", f.read(),
                        file_name=os.path.basename(out),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="excel_dl_btn")
            except Exception as e:
                st.error("Excel出力でエラーが発生しました。")
                st.exception(e)
