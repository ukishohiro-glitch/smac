# -*- coding: utf-8 -*-
"""
main.py（フル実装・外部定型文/ダミー一切なし）
- 単価参照は **master（原反価格シート）限定**。CSV の参照・切替 UI は排除。
- シート名は『原反価格』『OP』を最優先で解決（記号差・全半角差に耐性）。
- 得意先は『社名/支店/営業所』を**個別に**プルダウン化（存在するマスタだけ利用）。無い要素は自動で手入力にフォールバック。※『得意先一覧』単一シートにも対応（同シートから支店/営業所を自動派生）。
- UI フォントは 見出し=11pt／その他=9pt に統一。アプリ題名は『お見積書作成システム』。
- S・MAC見積に加えて、**部材（任意品）入力**セクションを追加（行追加、数量×単価で小計、サマリ/Excelに反映）。
- **部材は master の各シートを自動スキャン**し、『品名』『単価』（任意で『単位』『備考』）を持つシートを**部材カタログ**として参照可能に。
- CLI テストは同梱（本運用仕様の安全網）。

起動方法：
- GUI： `pip install streamlit pandas openpyxl` → `streamlit run main.py`
- CLI： `python main.py`
"""

import os, math, re, unicodedata
from io import BytesIO
from datetime import datetime
import pandas as pd

# --- Streamlit の有無 ---
HAS_STREAMLIT = True
try:
    import streamlit as st  # type: ignore
except Exception:
    HAS_STREAMLIT = False
    st = None

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# =========================
# ユーティリティ
# =========================
NUM_RE = re.compile(r"[-+]?\d[\d,]*\.?\d*")

def normalize_string(s):
    if s is None:
        return ""
    if isinstance(s, str):
        s = unicodedata.normalize("NFKC", s).replace("\u3000", " ")
        s = re.sub(r"\s+", " ", s.strip())
    return s

def parse_float(x):
    if x is None:
        return None
    if isinstance(x, (int, float)):
        return float(x)
    m = NUM_RE.search(str(x))
    return float(m.group()) if m else None

def extract_thickness(s):
    if s is None:
        return None
    m = re.search(r"(\d+(?:\.\d+)?)\s*(?:t|mm)", str(s).lower())
    return float(m.group(1)) if m else None

def ceil100(x):
    return int(math.ceil(float(x) / 100.0) * 100)

def pick_col(df: pd.DataFrame, candidates):
    if df is None or df.empty:
        return None
    cols = [normalize_string(c) for c in df.columns]
    df.columns = cols
    for c in candidates:
        if c in cols:
            return c
    for col in cols:
        for c in candidates:
            if str(c).lower() in col.lower():
                return col
    return None

# 共通エラー
class HaltError(RuntimeError):
    pass

def err_stop(msg: str):
    if HAS_STREAMLIT and st is not None:
        st.error(msg)
        st.stop()
    raise HaltError(msg)

# =========================
# 必須ファイル
# =========================
REQ_MASTER = "master.xlsx"
REQ_OP     = "OPマスタ.xlsx"  # 任意

def must_exist(path: str, label: str):
    if not os.path.exists(path):
        err_stop(f"必須ファイルが見つかりません: {label} → {path}")

# =========================
# マスタ読み込み（UI）
# =========================
if HAS_STREAMLIT:
    def _norm_sheet(s: str) -> str:
        s = normalize_string(s)
        s = re.sub(r"[\s・\-_/\.]+", "", s)
        return s.upper()

    def _resolve_sheet(xls: pd.ExcelFile, candidates: list[str]) -> str:
        target = {_norm_sheet(c): c for c in candidates}
        book = {_norm_sheet(n): n for n in xls.sheet_names}
        for k in target:
            if k in book:
                return book[k]
        raise HaltError("必要なシートが見つかりません: " + ", ".join(candidates))

    def _try_resolve_sheet(xls: pd.ExcelFile, candidates: list[str]):
        try:
            return _resolve_sheet(xls, candidates)
        except HaltError:
            return None

    @st.cache_data(show_spinner=False)
    def load_all():
        must_exist(REQ_MASTER, "master.xlsx")
        xls = pd.ExcelFile(REQ_MASTER)

        # シート候補（『原反価格』『OP』最優先）
        sheet_gen_candidates    = ["原反価格", "カーテン", "SMAC原反", "SMAC原反マスタ", "S MAC原反", "原反マスタ", "原反"]
        sheet_op_candidates     = ["OP", "SMAC-OP", "SMAC_OP", "SMACOP", "OPマスタ"]
        sheet_cat_candidates    = ["カーテン", "カーテンマスタ"]  # 任意
        sheet_perf_candidates   = ["カーテン性能", "性能", "性能マスタ"]  # 任意
        # ★ 得意先は『得意先一覧』にも対応
        sheet_cus_candidates    = ["得意先一覧", "得意先", "顧客", "取引先"]  # 任意
        sheet_branch_candidates = ["支店", "部署", "部支店"]  # 任意
        sheet_office_candidates = ["営業所", "事業所", "オフィス"]  # 任意

        sh_gen    = _resolve_sheet(xls, sheet_gen_candidates)
        sh_op     = _resolve_sheet(xls, sheet_op_candidates)
        sh_cat    = _try_resolve_sheet(xls, sheet_cat_candidates)
        sh_perf   = _try_resolve_sheet(xls, sheet_perf_candidates)
        sh_cus    = _try_resolve_sheet(xls, sheet_cus_candidates)
        sh_branch = _try_resolve_sheet(xls, sheet_branch_candidates)
        sh_office = _try_resolve_sheet(xls, sheet_office_candidates)

        def rd(name):
            return pd.read_excel(REQ_MASTER, sheet_name=name) if name else pd.DataFrame()

        df_gen     = rd(sh_gen)
        df_op      = rd(sh_op)
        df_catalog = rd(sh_cat)
        df_perf    = rd(sh_perf)
        df_cus     = rd(sh_cus)
        df_branch  = rd(sh_branch)
        df_office  = rd(sh_office)

        # 必須列の検証（同義語を許容）
        def _ensure(df, groups, label):
            miss = []
            for g in groups:
                if pick_col(df, g) is None:
                    miss.append(g[0])
            if miss:
                err_stop(f"シート『{label}』に必須列がありません: {miss}")
        _ensure(df_gen,
                [["製品名","品名","名称"],
                 ["原反幅(mm)","原反幅","幅","巾"],
                 ["厚み","厚さ","t"],
                 ["単価","単価(円/m)","原反単価","価格","金額"]],
                sh_gen)
        _ensure(df_op, [["OP名称"],["金額","単価"],["方向"]], sh_op)
        if not df_catalog.empty:
            _ensure(df_catalog, [["大分類"],["中分類"]], sh_cat)
        if not df_perf.empty:
            _ensure(df_perf, [["中分類"],["性能"]], sh_perf)
        if not df_cus.empty:
            # 得意先一覧の最小要件は社名（支店/営業所があれば後で派生に利用）
            _ensure(df_cus, [["社名"]], sh_cus)
        if not df_branch.empty:
            _ensure(df_branch, [["社名"],["支店名","支店"]], sh_branch)
        if not df_office.empty:
            _ensure(df_office, [["社名"],["支店名","支店"],["営業所名","営業所"]], sh_office)

        # master から単価抽出（標準名に正規化）
        name_col  = pick_col(df_gen, ["製品名","品名","名称"]) or df_gen.columns[0]
        price_col = pick_col(df_gen, ["単価","単価(円/m)","原反単価","価格","金額"])  # ここは必ず見つかる前提
        df_genprice_master = df_gen.rename(columns={name_col:"製品名", price_col:"単価"})[["製品名","単価"]]

        # ★『得意先一覧』のみの構成に対応：支店/営業所シートが無くても派生する
        if (df_branch.empty or df_office.empty) and not df_cus.empty:
            cus_c = pick_col(df_cus, ["社名"]) or df_cus.columns[0]
            br_c  = pick_col(df_cus, ["支店名","支店"])  # 任意
            of_c  = pick_col(df_cus, ["営業所名","営業所"])  # 任意
            if df_branch.empty and br_c:
                df_branch = df_cus[[cus_c, br_c]].dropna().drop_duplicates()
                df_branch.columns = ["社名","支店名"]
            if df_office.empty and br_c and of_c:
                df_office = df_cus[[cus_c, br_c, of_c]].dropna().drop_duplicates()
                df_office.columns = ["社名","支店名","営業所名"]

        # ★ 部材カタログ：既知以外の全シートをスキャンし、『品名』『単価』などを持つ場合に採用
        known = set([n for n in [sh_gen, sh_op, sh_cat, sh_perf, sh_cus, sh_branch, sh_office] if n])
        materials_catalogs: dict[str, pd.DataFrame] = {}
        for sh in xls.sheet_names:
            if sh in known:
                continue
            try:
                df_tmp = pd.read_excel(REQ_MASTER, sheet_name=sh)
            except Exception:
                continue
            if df_tmp is None or df_tmp.empty:
                continue
            nm = pick_col(df_tmp, ["品名","名称","品目"]) 
            pr = pick_col(df_tmp, ["単価","価格","金額"])
            if nm and pr:
                un = pick_col(df_tmp, ["単位"]) or None
                no = pick_col(df_tmp, ["備考","注記"]) or None
                cols_map = {nm:"品名", pr:"単価"}
                if un: cols_map[un] = "単位"
                if no: cols_map[no] = "備考"
                df_norm = df_tmp.rename(columns=cols_map)[[c for c in ["品名","単価","単位","備考"] if c in cols_map.values()]]
                # 型整形（単価は数値化）
                if "単価" in df_norm.columns:
                    df_norm["単価"] = df_norm["単価"].map(lambda v: parse_float(v) or None)
                materials_catalogs[sh] = df_norm.dropna(subset=["品名"]).reset_index(drop=True)
        
        return {
            "df_gen": df_gen,
            "df_op": df_op,
            "df_catalog": df_catalog,
            "df_perf": df_perf,
            "df_cus": df_cus,
            "df_branch": df_branch,
            "df_office": df_office,
            "df_genprice_master": df_genprice_master,
            "materials_catalogs": materials_catalogs,
        }

# =========================
# 見積ロジック（S・MAC）
# =========================
HEM_UNIT_THIN   = 450
HEM_UNIT_THICK  = 550
SEAM_UNIT_PER_M = 300
CUTTING_BASE_3  = 2000
CUTTING_BASE_4  = 3000

name_g = wcol = thcol = None  # UI で設定


def smac_estimate(middle_name: str, open_method: str, W: int, H: int, cnt: int, picked_ops_rows: list[dict]):
    """S・MAC見積の中核計算。
    戻り: dict(ok, sell_one, sell_total, note_ops, breakdown)
    """
    res = {"ok": False, "msg": "", "sell_one": 0, "sell_total": 0, "note_ops": [], "breakdown": {}}

    if any(v is None for v in [name_g, wcol, thcol]):
        res["msg"] = "内部エラー: 列参照が未設定です。"; return res
    if 'df_gen_merged' not in globals() or 'df_op' not in globals():
        res["msg"] = "内部エラー: マスタが未設定です。"; return res
    if not middle_name or W<=0 or H<=0 or cnt<=0:
        res["msg"] = "中分類/寸法/数量を確認してください。"; return res

    hit = df_gen_merged[df_gen_merged[name_g]==middle_name]
    if hit.empty:
        res["msg"] = f"原反価格に『{middle_name}』が見つかりません。"; return res

    gen_width = parse_float(hit.iloc[0][wcol])
    gen_price = parse_float(hit.iloc[0]["単価"])  # 円/m
    thick = extract_thickness(hit.iloc[0][thcol])
    if not gen_width or not gen_price:
        res["msg"] = "原反幅または単価が不正です。"; return res

    # 寸法補正
    if open_method == "片引き":
        cur_w = W * 1.05; panels = 1
    else:
        cur_w = (W/2) * 1.05; panels = 2
    cur_h = H + 50

    # 原反使用量（1間口）
    length_per_panel_m = (cur_h * 1.2) / 1000.0
    joints = math.ceil(cur_w / gen_width)
    raw_len_m = length_per_panel_m * joints * panels
    raw_one = gen_price * raw_len_m

    # 裁断賃
    cutting_one = (CUTTING_BASE_3 if joints <= 3 else CUTTING_BASE_4) * panels

    # 幅繋ぎ
    seams_total = max(0, joints - 1) * panels
    seam_one = math.ceil(cur_h/1000.0) * SEAM_UNIT_PER_M * seams_total

    # 四方折り返し
    hem_unit = HEM_UNIT_THIN if (thick is not None and thick <= 0.3) else HEM_UNIT_THICK
    hem_perimeter_m = (cur_w + cur_w + cur_h + cur_h) / 1000.0
    fourfold_one = math.ceil(hem_perimeter_m) * hem_unit * panels

    # OP（方向 W/横/X→幅、その他→高さ）
    name_col = pick_col(df_op, ["OP名称"]) or "OP名称"
    price_col= pick_col(df_op, ["金額","単価"]) or "金額"
    dir_col  = pick_col(df_op, ["方向"]) or "方向"

    op_total = 0
    note_ops = []
    for row in (picked_ops_rows or []):
        nm = normalize_string(row.get(name_col, ""))
        if not nm: continue
        unit = int(parse_float(row.get(price_col)) or 0)
        dire = normalize_string(row.get(dir_col, "")).upper()
        if unit <= 0: continue
        base_mm = cur_w if dire in ["W","横","X"] else cur_h
        units_1000 = math.ceil(base_mm/1000.0)
        sub = units_1000 * unit * panels * cnt
        op_total += sub
        note_ops.append(nm)

    op_one = op_total / cnt if cnt else 0

    genka_one   = raw_one + cutting_one + seam_one + fourfold_one + op_one
    genka_total = raw_one*cnt + cutting_one*cnt + seam_one*cnt + fourfold_one*cnt + op_total
    sell_one    = ceil100(genka_one / 0.6)
    sell_total  = ceil100(genka_total / 0.6)

    res.update({
        "ok": True,
        "sell_one": int(sell_one),
        "sell_total": int(sell_total),
        "note_ops": note_ops,
        "breakdown": {
            "原反使用量(m)":       round(raw_len_m, 2),
            "原反幅(mm)":          int(round(gen_width)),
            "原反単価(円/m)":       int(round(gen_price)),
            "原反材料(1式)":        int(round(raw_one)),
            "裁断賃(1式)":          int(round(cutting_one)),
            "幅繋ぎ(1式)":          int(round(seam_one)),
            "四方折り返し(1式)":    int(round(fourfold_one)),
            "OP加算(1式)":          int(round(op_one)),
            "原価(1式)":            int(round(genka_one)),
            "販売単価(1式)":        int(sell_one),
            "販売金額(数量分)":      int(sell_total),
            "粗利率":               (max(0.0, 1.0 - (genka_total / sell_total)) if sell_total else 0.0),
        }
    })
    return res

# =========================
# 部材（任意品）ロジック
# =========================

def build_material_items(rows: list[dict]):
    """部材行から明細（品名/数量/単位/単価/小計/備考/種別）を生成。戻り: (items, total)"""
    items = []
    total = 0
    for r in rows:
        name = normalize_string(r.get("品名", ""))
        if not name:
            continue
        qty  = int(parse_float(r.get("数量")) or 0)
        unit = normalize_string(r.get("単位", "")) or "個"
        price= int(parse_float(r.get("単価")) or 0)
        note = normalize_string(r.get("備考", ""))
        if qty <= 0 or price <= 0:
            continue
        sub = qty * price
        items.append({
            "品名": name,
            "数量": qty,
            "単位": unit,
            "単価": price,
            "小計": sub,
            "種別": "部材",
            "備考": note,
        })
        total += sub
    return items, total

# =========================
# 見積 Excel（明細）
# =========================

def is_opening_head(it: dict) -> bool:
    name = str(it.get("品名",""))
    kind = str(it.get("種別",""))
    return ("S・MAC" in kind) or ("S・MAC" in name)


def extract_mark(note: str) -> str | None:
    if not note:
        return None
    m = re.search(r"符号[:：]\s*([^／\s]+)", str(note))
    return m.group(1) if m else None


def build_estimate_workbook(header: dict, items: list[dict]) -> BytesIO:
    wb = Workbook(); ws = wb.active; ws.title = "お見積書（明細）"
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor="F2F2F2")
    yen_fmt = '"¥"#,##0'; int_fmt = '#,##0'
    widths = [36, 8, 7, 11, 12, 12, 56]
    for i,w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    r = 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    ws.cell(r,1,"お見積書（明細）").font = Font(size=11, bold=True); r+=2
    ws.cell(r,1,"見積番号").font=Font(size=9); ws.cell(r,2, header.get("estimate_no",""))
    ws.cell(r,4,"作成日").font=Font(size=9);   ws.cell(r,5, header.get("date","")); r+=1
    ws.cell(r,1,"得意先").font=Font(size=9); ws.cell(r,2, header.get("customer_name",""))
    ws.cell(r,4,"部署・支店").font=Font(size=9); ws.cell(r,5, header.get("branch_name","")); r+=1
    ws.cell(r,1,"事業所").font=Font(size=9); ws.cell(r,2, header.get("office_name",""))
    ws.cell(r,4,"ご担当").font=Font(size=9); ws.cell(r,5, header.get("person_name","")); r+=2
    hdr = ["品名","数量","単位","単価","小計","種別","備考"]
    for c,t in enumerate(hdr, start=1):
        cell = ws.cell(r,c,t)
        cell.font = Font(size=9, bold=True); cell.fill = head_fill
        cell.alignment = Alignment(horizontal="center"); cell.border = border
    r+=1
    total = 0; prev_mark=None; started=False
    for it in items:
        mark = extract_mark(it.get("備考","")); is_head = is_opening_head(it)
        if started and ((mark and mark!=prev_mark) or is_head):
            r += 1
        prev_mark = mark if mark else prev_mark; started=True
        row_vals = [
            it.get("品名",""),
            it.get("数量","") if it.get("数量","")!="" else None,
            it.get("単位",""),
            it.get("単価","") if it.get("単価","")!="" else None,
            it.get("小計","") if it.get("小計","")!="" else None,
            it.get("種別",""),
            it.get("備考",""),
        ]
        for c,v in enumerate(row_vals, start=1):
            cell=ws.cell(r,c,v); cell.border=border; cell.font=Font(size=9)
            if c in (2,4,5) and isinstance(v,(int,float)):
                cell.number_format = int_fmt if c==2 else yen_fmt
            if c==2: cell.alignment=Alignment(horizontal="right")
            if c in (4,5): cell.alignment=Alignment(horizontal="right")
        if isinstance(it.get("小計"), (int,float)):
            total += it["小計"]
        r+=1
    r+=1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.cell(r,1,"合計").font=Font(size=9, bold=True)
    ws.cell(r,5,total).number_format=yen_fmt; ws.cell(r,5).font=Font(size=9, bold=True)
    for c in (1,5): ws.cell(r,c).border=border
    bio=BytesIO(); wb.save(bio); bio.seek(0); return bio

# =========================
# UI（Streamlit）
# =========================
if HAS_STREAMLIT:
    st.set_page_config(page_title="お見積書作成システム", layout="wide")
    # グローバル CSS（見出し=11pt／その他=9pt）
    st.markdown(
        """
        <style>
        html, body, [data-testid="stAppViewContainer"] * { font-size: 9pt !important; }
        h1, h2, h3, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 { font-size: 11pt !important; }
        [data-testid="stMetricValue"], [data-testid="stMetricLabel"] { font-size: 9pt !important; }
        thead, tbody, .stDataFrame, .stTable { font-size: 9pt !important; }
        .stButton>button, .stDownloadButton>button, .stTextInput input, .stNumberInput input,
        .stSelectbox div, .stRadio div { font-size: 9pt !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title("お見積書作成システム")

    M = load_all()
    # 列名を正規化
    for k in ["df_gen","df_op","df_catalog","df_perf","df_cus","df_branch","df_office","df_genprice_master"]:
        df = M.get(k, pd.DataFrame())
        if isinstance(df, pd.DataFrame) and not df.empty:
            df.columns = [normalize_string(c) for c in df.columns]
        M[k] = df

    df_gen      = M["df_gen"].copy()
    df_op       = M["df_op"].copy()
    df_catalog  = M["df_catalog"].copy()
    df_perf     = M["df_perf"].copy()
    df_cus      = M["df_cus"].copy()
    df_branch   = M["df_branch"].copy()
    df_office   = M["df_office"].copy()
    df_genprice = M["df_genprice_master"].copy()  # master 限定
    materials_catalogs: dict[str, pd.DataFrame] = M.get("materials_catalogs", {})

    # 原反の単価結合（製品名で左結合）
    name_g = pick_col(df_gen, ["製品名","品名","名称"]) or df_gen.columns[0]
    wcol   = pick_col(df_gen, ["原反幅(mm)","原反幅","幅","巾"]) or df_gen.columns[1]
    thcol  = pick_col(df_gen, ["厚み","厚さ","t"]) or name_g

    pname_csv = pick_col(df_genprice, ["製品名","品名","名称","中分類"]) or "製品名"
    pcol_csv  = pick_col(df_genprice, ["単価","単価(円/m)","原反単価","価格","金額"]) or "単価"

    pcol_gen = pick_col(df_gen, ["単価","単価(円/m)","原反単価","価格","金額"])  # 任意（衝突回避）
    _g = df_gen.copy()
    if pcol_gen:
        _g = _g.rename(columns={pcol_gen: f"{pcol_gen}_GEN"})

    dfp = df_genprice.rename(columns={pname_csv: "製品名", pcol_csv: "単価"})[["製品名","単価"]]
    _df = pd.merge(_g, dfp, on="製品名", how="left")

    if "単価" not in _df.columns:
        err_stop("原反単価の結合に失敗しました（列名の衝突または未検出）。master の列名をご確認ください。")
    if _df["単価"].isna().any():
        miss = _df[_df["単価"].isna()]["製品名"].dropna().unique().tolist()
        err_stop("単価未登録の製品があります: " + ", ".join(map(str, miss[:10])) + (" ..." if len(miss)>10 else ""))

    df_gen_merged = _df

    # ---- 見積番号生成（UC37MM-####）
    def gen_estimate_no():
        today = datetime.today(); prefix = f"UC37{today.month:02d}-"
        if "_est_seq" not in st.session_state:
            st.session_state["_est_seq"] = 0
        st.session_state["_est_seq"] += 1
        return prefix + f"{st.session_state['_est_seq']%10000:04d}"

    # 共通：既存値に基づき index を算出する selectbox ヘルパ
    def _selectbox_or_text(label: str, options: list, key: str):
        options = options or []
        if options:
            cur = st.session_state.get(key)
            idx = options.index(cur) if (cur in options) else 0
            return st.selectbox(label, options, index=idx, key=key)
        else:
            return st.text_input(label, value=st.session_state.get(key, ""), key=key)

    with st.container():
        st.markdown("---")
        st.markdown("### 得意先情報入力")

        if "estimate_no" not in st.session_state:
            st.session_state["estimate_no"] = gen_estimate_no()
        if "created_disp" not in st.session_state:
            st.session_state["created_disp"] = datetime.today().strftime("%Y/%m/%d")

        col1, col2, col3 = st.columns([1.0, 1.0, 1.0])
        with col1:
            # 社名：マスタがあればプルダウン、無ければ手入力
            if not df_cus.empty:
                c_cus = pick_col(df_cus, ["社名"]) or df_cus.columns[0]
                cus_list = df_cus[c_cus].dropna().unique().tolist()
                customer = _selectbox_or_text("社名（得意先名）", cus_list, key="customer")
            else:
                customer = _selectbox_or_text("社名（得意先名）", [], key="customer")
        with col2:
            # 支店：支店マスタがあれば（社名フィルタ後を）プルダウン／無い場合は『得意先一覧』内の列から派生
            branches = []
            if not df_branch.empty:
                c_b_cus = pick_col(df_branch, ["社名"]) or df_branch.columns[0]
                c_b = pick_col(df_branch, ["支店名","支店"]) or df_branch.columns[1]
                branches = df_branch[df_branch[c_b_cus]==customer][c_b].dropna().unique().tolist() if customer else []
            elif not df_cus.empty:
                c_b = pick_col(df_cus, ["支店名","支店"])  # 任意
                c_c = pick_col(df_cus, ["社名"]) or df_cus.columns[0]
                if c_b:
                    branches = df_cus[df_cus[c_c]==customer][c_b].dropna().unique().tolist() if customer else []
            branch = _selectbox_or_text("支店名", branches, key="branch")
        with col3:
            # 営業所：営業所マスタがあれば（社名・支店でフィルタ後を）プルダウン／無い場合は『得意先一覧』内から派生
            offices = []
            if not df_office.empty:
                c_o_cus = pick_col(df_office, ["社名"]) or df_office.columns[0]
                c_o_b = pick_col(df_office, ["支店名","支店"]) or df_office.columns[1]
                c_o = pick_col(df_office, ["営業所名","営業所"]) or df_office.columns[2]
                offices = df_office[(df_office[c_o_cus]==customer) & (df_office[c_o_b]==branch)][c_o].dropna().unique().tolist() if (customer and branch) else []
            elif not df_cus.empty:
                c_o = pick_col(df_cus, ["営業所名","営業所"])  # 任意
                c_b = pick_col(df_cus, ["支店名","支店"])  # 任意
                c_c = pick_col(df_cus, ["社名"]) or df_cus.columns[0]
                if c_o and c_b:
                    offices = df_cus[(df_cus[c_c]==customer) & (df_cus[c_b]==branch)][c_o].dropna().unique().tolist() if (customer and branch) else []
            office = _selectbox_or_text("営業所名", offices, key="office")

        col4, col5, col6 = st.columns([0.8, 0.6, 1.4])
        with col4:
            est = st.text_input("見積番号", value=st.session_state["estimate_no"]); st.session_state["estimate_no"] = est
        with col5:
            created = st.text_input("作成日（YYYY/MM/DD）", value=st.session_state["created_disp"]); st.session_state["created_disp"] = created
        with col6:
            project = st.text_input("物件名", value=st.session_state.get("project_name","")); st.session_state["project_name"] = project

        st.caption(
            f"見積番号：{st.session_state['estimate_no']} / 作成日：{st.session_state['created_disp']} / "
            f"得意先：{customer or '（未入力）'} / 支店：{branch or '（未入力）'} / 営業所：{office or '（未入力）'} / 物件名：{project or '（未入力）'}"
        )
        st.caption("単価参照元：master（原反価格）")
        st.markdown("---")

    if "openings" not in st.session_state:
        st.session_state.openings = [1]

    overall_items = []
    overall_total = 0

    def add_total(v):
        globals()["overall_total"] = globals().get("overall_total", 0) + int(v or 0)

    # --- 間口（S・MAC） ---
    def render_opening(idx: int):
        pref = f"o{idx}_"

        a1,a2,a3,a4 = st.columns([0.7,1.0,1.0,0.7])
        with a1: mark = st.text_input("符号", key=pref+"mark")
        with a2: W = st.number_input("間口W (mm)", min_value=0, value=0, step=50, key=pref+"w")
        with a3: H = st.number_input("間口H (mm)", min_value=0, value=0, step=50, key=pref+"h")
        with a4: CNT = st.number_input("数量", min_value=1, value=1, step=1, key=pref+"cnt")

        st.markdown("#### カーテン（S・MAC）")
        mids = df_gen_merged[name_g].dropna().unique().tolist()
        mid = st.selectbox("中分類（原反・製品名）", [""]+mids, key=pref+"mid")

        perf_opts = []
        if not M["df_perf"].empty:
            perf_mid_col  = pick_col(M["df_perf"], ["中分類"]) or M["df_perf"].columns[0]
            perf_perf_col = pick_col(M["df_perf"], ["性能"]) or M["df_perf"].columns[1]
            perf_opts = M["df_perf"][M["df_perf"][perf_mid_col]==mid][perf_perf_col].dropna().unique().tolist() if mid else []
        _ = st.selectbox("カーテン性能（任意）", [""]+perf_opts, key=pref+"perf")

        open_mtd = st.radio("片引き/引分け", ["片引き","引分け"], key=pref+"open", horizontal=True)

        name_col = pick_col(df_op, ["OP名称"]) or "OP名称"
        price_col= pick_col(df_op, ["金額","単価"]) or "金額"
        dir_col  = pick_col(df_op, ["方向"]) or "方向"

        op_rows = []
        if not df_op.empty and all(c in df_op.columns for c in [name_col, price_col, dir_col]):
            names = df_op[name_col].dropna().unique().tolist()
            cols = st.columns(3)
            for i, nm in enumerate(names):
                with cols[i%3]:
                    if st.checkbox(nm, key=pref+f"op_{i}"):
                        row = df_op[df_op[name_col]==nm].iloc[0].to_dict()
                        op_rows.append(row)

        if W>0 and H>0 and CNT>0 and mid:
            r = smac_estimate(mid, open_mtd, W, H, CNT, op_rows)
            if r["ok"]:
                note = f"W{W}×H{H}mm"
                if r["note_ops"]: note += "／OP:" + "・".join(r["note_ops"])
                overall_items.append({
                    "品名": f"S・MACカーテン {mid} {open_mtd}",
                    "数量": CNT, "単位":"式",
                    "単価": r["sell_one"], "小計": r["sell_total"],
                    "種別": "S・MAC",
                    "備考": (f"符号:{mark}／" if mark else "") + note,
                })
                add_total(r["sell_total"])

                bd = r.get("breakdown", {})
                if bd:
                    with st.expander("原価構成（1間口あたり）", expanded=False):
                        order = [
                            "原反使用量(m)", "原反幅(mm)", "原反単価(円/m)",
                            "原反材料(1式)", "裁断賃(1式)", "幅繋ぎ(1式)", "四方折り返し(1式)", "OP加算(1式)",
                            "原価(1式)", "販売単価(1式)", "販売金額(数量分)", "粗利率",
                        ]
                        rows = []
                        for k in order:
                            if k not in bd: continue
                            v = bd[k]
                            if k == "原反使用量(m)": rows.append([k, f"{float(v):.2f} m"])
                            elif k == "原反幅(mm)": rows.append([k, f"{int(v)} mm"])
                            elif k == "原反単価(円/m)": rows.append([k, f"¥{int(v):,}/m"])
                            elif k == "粗利率": rows.append([k, f"{float(v)*100:.1f}%"])
                            else: rows.append([k, f"¥{int(v):,}"])
                        st.dataframe(pd.DataFrame(rows, columns=["項目","金額"]), use_container_width=True, hide_index=True)
            else:
                st.warning(r.get("msg") or "S・MACの計算に失敗しました。")

    st.markdown("---")
    col_a, col_b = st.columns([0.7,0.3])
    with col_b:
        if st.button("＋ 間口を追加", key="add_opening_btn"):
            st.session_state.openings.append(len(st.session_state.openings)+1)
            st.rerun()
    with col_a:
        for i,_ in enumerate(st.session_state.openings, start=1):
            st.markdown(f"#### 間口 {i}")
            render_opening(i)

    # --- 部材（任意品）入力 ---
    st.markdown("---")
    st.markdown("### 部材（任意品）入力")

    # 1) 部材カタログから追加（master の各シート参照）
    if materials_catalogs:
        sh_names = sorted(list(materials_catalogs.keys()))
        colm1, colm2, colm3, colm4, colm5 = st.columns([1.2, 1.2, 0.7, 0.8, 0.6])
        with colm1:
            sel_sheet = st.selectbox("参照シート", sh_names, key="mat_src_sheet")
        df_src = materials_catalogs.get(sel_sheet, pd.DataFrame())
        with colm2:
            item_names = df_src["品名"].dropna().unique().tolist() if not df_src.empty and "品名" in df_src.columns else []
            sel_item = st.selectbox("品名（カタログ）", item_names, key="mat_src_item") if item_names else None
        # 既定値の抽出
        base_unit = ""
        base_price = None
        base_note = ""
        if sel_item and not df_src.empty:
            row = df_src[df_src["品名"]==sel_item].iloc[0]
            base_unit = str(row.get("単位", "") or "")
            base_price = int(parse_float(row.get("単価")) or 0)
            base_note = str(row.get("備考", "") or "")
        with colm3:
            qty = st.number_input("数量", min_value=1, value=1, step=1, key="mat_src_qty")
        with colm4:
            unit = st.text_input("単位", value=base_unit, key="mat_src_unit")
        with colm5:
            price = st.number_input("単価", min_value=0, value=(base_price or 0), step=10, key="mat_src_price")
        note = st.text_input("備考", value=base_note, key="mat_src_note")
        if st.button("この部材を追加", key="mat_add_btn") and sel_item:
            new_row = {"品名": sel_item, "数量": qty, "単位": unit, "単価": int(price), "備考": note}
            if "materials_df" not in st.session_state or st.session_state.materials_df is None:
                st.session_state.materials_df = pd.DataFrame([new_row])
            else:
                st.session_state.materials_df = pd.concat([st.session_state.materials_df, pd.DataFrame([new_row])], ignore_index=True)
            st.success("部材を追加しました。下の表で編集・削除できます。")

    # 2) 直接編集（行追加・削除可）
    cols = ["品名","数量","単位","単価","備考"]
    if "materials_df" not in st.session_state:
        st.session_state.materials_df = pd.DataFrame([{c:("" if c!="数量" else 1) for c in cols}])
    df_mats = st.data_editor(
        st.session_state.materials_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="materials_editor",
    )
    st.session_state.materials_df = df_mats
    mats_items, mats_total = build_material_items(df_mats.to_dict(orient="records"))

    if mats_items:
        overall_items.extend(mats_items)
        add_total(mats_total)

    # サマリ
    st.markdown("---")
    st.markdown("### 見積サマリ")
    if overall_items:
        cols_o = ["品名","数量","単位","単価","小計","種別","備考"]
        rows = []
        prev_mark = None; started = False
        for it in overall_items:
            mark = extract_mark(it.get("備考","")); head = is_opening_head(it)
            if started and ((mark and mark!=prev_mark) or head):
                rows.append({c: "" for c in cols_o})
            prev_mark = mark if mark else prev_mark; started=True
            rows.append({c: it.get(c, "") for c in cols_o})
        df_sum = pd.DataFrame(rows, columns=cols_o)
        def _n(v):
            if v in (None, "", "-"): return ""
            try: return f"{int(v):,}"
            except: return v
        for c in ["数量","単価","小計"]:
            df_sum[c] = df_sum[c].map(_n)
        st.dataframe(df_sum, use_container_width=True, hide_index=True)
    else:
        st.info("明細がありません。")

    st.markdown("### 見積金額")
    st.metric("税抜合計", f"¥{globals().get('overall_total', 0):,}")

    # Excel
    st.markdown("---")
    st.markdown("### Excel出力")
    if st.button("Excel保存", key="excel_save_btn_v1"):
        header = {
            "estimate_no": st.session_state.get("estimate_no",""),
            "date":        st.session_state.get("created_disp", datetime.today().strftime("%Y/%m/%d")),
            "customer_name": st.session_state.get("customer",""),
            "branch_name":   st.session_state.get("branch",""),
            "office_name":   st.session_state.get("office",""),
            "person_name":   st.session_state.get("pic",""),
        }
        try:
            bio = build_estimate_workbook(header, overall_items)
            st.success("Excelを作成しました。")
            st.download_button(
                "ダウンロード（お見積書_明細.xlsx）",
                data=bio.getvalue(),
                file_name="お見積書_明細.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="dl_xlsx_v1",
            )
        except Exception as e:
            st.error("Excel出力でエラーが発生しました。")
            st.exception(e)

# =========================
# CLI テスト
# =========================
else:
    def _fixture_frames():
        df_gen = pd.DataFrame({
            "製品名": ["標準0.25t", "標準0.4t"],
            "原反幅(mm)": [1370, 1500],
            "厚み": ["0.25t", "0.4t"],
            "単価": [1800, 2200],
        })
        df_op = pd.DataFrame({
            "OP名称": ["透明窓追加", "補強テープ", "H方向OP"],
            "金額": [1000, 800, 1000],
            "方向": ["W", "W", "H"],
        })
        return df_gen, df_op

    def _prepare_globals_for_tests():
        global df_gen_merged, df_op, name_g, wcol, thcol
        df_gen, df_op_local = _fixture_frames()
        df_op = df_op_local
        name_g = "製品名"; wcol = "原反幅(mm)"; thcol = "厚み"
        df_gen_merged = df_gen  # すでに単価を含む

    def run_tests():
        _prepare_globals_for_tests()

        # Test 1: 基本計算（片引き）
        r1 = smac_estimate("標準0.25t", "片引き", W=2000, H=2000, cnt=1, picked_ops_rows=[])
        assert r1["ok"] is True and r1["sell_total"] > 0
        for k in ["原反使用量(m)", "原反幅(mm)", "原反単価(円/m)"]:
            assert k in r1["breakdown"]

        # Test 2: 引分けは片引き以上
        r2 = smac_estimate("標準0.25t", "引分け", W=2000, H=2000, cnt=1, picked_ops_rows=[])
        assert r2["ok"] is True and r2["sell_total"] >= r1["sell_total"]

        # Test 3: OP方向差（Hの方が高い想定）
        rW = smac_estimate("標準0.4t", "片引き", W=1500, H=2500, cnt=1,
                           picked_ops_rows=[{"OP名称":"透明窓追加","金額":1000,"方向":"W"}])
        rH = smac_estimate("標準0.4t", "片引き", W=1500, H=2500, cnt=1,
                           picked_ops_rows=[{"OP名称":"H方向OP","金額":1000,"方向":"H"}])
        assert rW["ok"] and rH["ok"] and rH["sell_total"] >= rW["sell_total"]

        # Test 4: 不正入力
        r_bad = smac_estimate("標準0.25t", "片引き", W=0, H=1000, cnt=1, picked_ops_rows=[])
        assert r_bad["ok"] is False

        # Test 5: 100円切上げ
        assert ceil100(1) == 100 and ceil100(100) == 100 and ceil100(101) == 200
        assert r1["sell_one"] % 100 == 0

        # Test 6: 厚みが厚い方が折返し高い
        r_thin = smac_estimate("標準0.25t", "片引き", W=1800, H=2000, cnt=1, picked_ops_rows=[])
        r_thick= smac_estimate("標準0.4t",  "片引き", W=1800, H=2000, cnt=1, picked_ops_rows=[])
        assert r_thin["ok"] and r_thick["ok"]
        assert r_thick["breakdown"]["四方折り返し(1式)"] >= r_thin["breakdown"]["四方折り返し(1式)"]

        # Test 7: 部材アイテム生成
        mats, tot = build_material_items([
            {"品名":"ブラケット", "数量":2, "単位":"個", "単価":350, "備考":"--"},
            {"品名":"ビス", "数量":10, "単位":"本", "単価":15},
            {"品名":"空行", "数量":0, "単価":100},  # 無視されるべき
        ])
        assert len(mats) == 2 and tot == 2*350 + 10*15

        print("All tests passed.")

    if __name__ == "__main__":
        try:
            run_tests()
        except AssertionError as e:
            print("TEST FAILED:", e)
        except Exception as e:
            print("ERROR:", e)
