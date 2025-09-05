# -*- coding: utf-8 -*-
"""
main.py（フル実装・外部定型文/ダミーなし） — シート名「原反価格」「OP」に対応 & 得意先系シートが無くてもUIで手入力に自動フォールバック

今回の修正点（エラー解消）：
- 画像の例外は、得意先系シート（得意先/支店/営業所）が存在しないにもかかわらず **必須扱い** で解決を試み、`HaltError` を投げていたことが原因。
- これらのシートを **任意扱い** に変更し、見つからない場合は **テキスト入力に自動フォールバック** するように修正。
- `_try_resolve_sheet` を追加して、任意シートは見つからなければ `None` を返す設計に変更。
- 既存テストは維持、`sys.exit` は不使用のまま。

使い方：
- GUI（推奨）：`pip install streamlit pandas openpyxl` → `streamlit run main.py`
- CLIテスト（streamlit無し可）：`python main.py`
"""

import os, math, re, unicodedata
from io import BytesIO
from datetime import datetime
import pandas as pd

# --- Streamlitの有無を検出 ---
HAS_STREAMLIT = True
try:
    import streamlit as st  # type: ignore
except Exception:
    HAS_STREAMLIT = False
    st = None  # sentinel

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
        s = unicodedata.normalize("NFKC", s)
        s = s.replace("\u3000", " ")
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
    # 部分一致でも許容
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
REQ_OP     = "OPマスタ.xlsx"   # 任意
REQ_GENCSV = "原反単価表.csv"

def must_exist(path: str, label: str):
    if not os.path.exists(path):
        err_stop(f"必須ファイルが見つかりません: {label} → {path}")

# =========================
# マスタ読み込み（UIモード専用）
# =========================
if HAS_STREAMLIT:
    # シート名のゆるい解決: 記号（・ - _ / . 空白）差と大/小/全/半角差を吸収
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

    # 見つからない場合は None を返す（任意シート向け）
    def _try_resolve_sheet(xls: pd.ExcelFile, candidates: list[str]):
        try:
            return _resolve_sheet(xls, candidates)
        except HaltError:
            return None

    @st.cache_data(show_spinner=False)
    def load_all():
        must_exist(REQ_MASTER, "master.xlsx")
        must_exist(REQ_GENCSV, "原反単価表.csv")

        xls = pd.ExcelFile(REQ_MASTER)

        # ★ユーザー指定を最優先（原反=原反価格、OP=OP）。記号差にも強い候補一覧。
        sheet_gen_candidates    = ["原反価格", "カーテン", "SMAC原反", "SMAC原反マスタ", "S MAC原反", "原反マスタ", "原反"]
        sheet_op_candidates     = ["OP", "SMAC-OP", "SMAC_OP", "SMACOP", "OPマスタ"]
        # 以下は任意
        sheet_cat_candidates    = ["カーテン", "カーテンマスタ"]
        sheet_perf_candidates   = ["カーテン性能", "性能", "性能マスタ"]
        sheet_cus_candidates    = ["得意先", "顧客", "取引先"]
        sheet_branch_candidates = ["支店", "部署", "部支店"]
        sheet_office_candidates = ["営業所", "事業所", "オフィス"]

        # 必須シート
        sh_gen    = _resolve_sheet(xls, sheet_gen_candidates)
        sh_op     = _resolve_sheet(xls, sheet_op_candidates)
        # 任意シート（無ければ None）
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

        # 検証（必須のみ）
        for name, df, must in [
            (sh_gen,    df_gen,     ["製品名","原反幅(mm)","厚み"]),
            (sh_op,     df_op,      ["OP名称","金額","方向"]),
        ]:
            if df.empty:
                err_stop(f"master.xlsx のシート『{name}』が見つからないか空です。")
            miss = [c for c in must if pick_col(df, [c]) is None]
            if miss:
                err_stop(f"シート『{name}』に必須列がありません: {miss}")

        # 任意シート（存在すれば列チェック）
        for name, df, must in [
            (sh_cat,    df_catalog, ["大分類","中分類"]),
            (sh_perf,   df_perf,    ["中分類","性能"]),
            (sh_cus,    df_cus,     ["社名"]),
            (sh_branch, df_branch,  ["社名","支店名"]),
            (sh_office, df_office,  ["社名","支店名","営業所名"]),
        ]:
            if name and not df.empty:
                miss = [c for c in must if pick_col(df, [c]) is None]
                if miss:
                    err_stop(f"シート『{name}』に必須列がありません: {miss}")

        # 原反単価（CSV）
        df_genprice = pd.read_csv(REQ_GENCSV, encoding="utf-8")
        if not set(["製品名","単価"]).issubset(df_genprice.columns):
            err_stop("原反単価表.csv に必須列 '製品名','単価' がありません。")

        # OPマスタ（任意・優先）
        if os.path.exists(REQ_OP):
            xls_op = pd.ExcelFile(REQ_OP)
            try:
                sh_op2 = _resolve_sheet(xls_op, sheet_op_candidates)
                df_op2 = pd.read_excel(REQ_OP, sheet_name=sh_op2)
                if not df_op2.empty:
                    name_col = pick_col(df_op2, ["OP名称"]) or df_op2.columns[0]
                    df_tmp = pd.concat([df_op2, df_op], ignore_index=True)
                    if name_col in df_tmp.columns:
                        df_op = df_tmp.drop_duplicates(subset=[name_col], keep="first")
            except HaltError:
                pass

        return {
            "df_gen": df_gen,
            "df_op": df_op,
            "df_catalog": df_catalog,
            "df_perf": df_perf,
            "df_cus": df_cus,
            "df_branch": df_branch,
            "df_office": df_office,
            "df_genprice": df_genprice,
        }

# =========================
# S・MAC 計算（仕様準拠）
# =========================
HEM_UNIT_THIN   = 450     # 四方折り返し（~0.3t）
HEM_UNIT_THICK  = 550     # 四方折り返し（0.3t~）
SEAM_UNIT_PER_M = 300     # 幅繋ぎ（1mあたり）
CUTTING_BASE_3  = 2000    # 裁断賃（継ぎ<=3）
CUTTING_BASE_4  = 3000    # 裁断賃（継ぎ>=4）

# 以降のグローバル参照は UI モードでセット／CLIテストではフィクスチャで上書き
name_g = wcol = thcol = None
# df_gen_merged, df_op は後でセット

def smac_estimate(middle_name: str, open_method: str, W: int, H: int, cnt: int, picked_ops_rows: list[dict]):
    """戻り: dict(ok, sell_one, sell_total, note_ops, breakdown)
    middle_name は 原反マスタ『製品名』と一致している前提。
    """
    res = {"ok": False, "msg": "", "sell_one": 0, "sell_total": 0, "note_ops": [], "breakdown": {}}

    # 必須グローバル検証
    if any(v is None for v in [name_g, wcol, thcol]):
        res["msg"] = "内部エラー: 列参照が未設定です。"
        return res
    if 'df_gen_merged' not in globals() or 'df_op' not in globals():
        res["msg"] = "内部エラー: マスタが未設定です。"
        return res

    if not middle_name or W<=0 or H<=0 or cnt<=0:
        res["msg"] = "中分類/寸法/数量を確認してください。"
        return res

    hit = df_gen_merged[df_gen_merged[name_g]==middle_name]
    if hit.empty:
        res["msg"] = f"原反価格に『{middle_name}』が見つかりません。"
        return res

    gen_width = parse_float(hit.iloc[0][wcol])
    gen_price = parse_float(hit.iloc[0]["単価"])  # 円/m
    thick = extract_thickness(hit.iloc[0][thcol])
    if not gen_width or not gen_price:
        res["msg"] = "原反幅または単価が不正です。"
        return res

    # 寸法補正
    if open_method == "片引き":
        cur_w = W * 1.05; panels = 1
    else:
        cur_w = (W/2) * 1.05; panels = 2
    cur_h = H + 50

    # 原反使用量（1間口）
    length_per_panel_m = (cur_h * 1.2) / 1000.0
    joints = math.ceil(cur_w / gen_width)              # 1パネル内の継ぎ数
    raw_len_m = length_per_panel_m * joints * panels   # 使用量[m]
    raw_one = gen_price * raw_len_m                    # 原反材料（1式）

    # 裁断賃
    cutting_one = (CUTTING_BASE_3 if joints <= 3 else CUTTING_BASE_4) * panels

    # 幅繋ぎ（縦継ぎ本数）
    seams_total = max(0, joints - 1) * panels
    seam_one = math.ceil(cur_h/1000.0) * SEAM_UNIT_PER_M * seams_total

    # 四方折り返し
    hem_unit = HEM_UNIT_THIN if (thick is not None and thick <= 0.3) else HEM_UNIT_THICK
    hem_perimeter_m = (cur_w + cur_w + cur_h + cur_h) / 1000.0
    fourfold_one = math.ceil(hem_perimeter_m) * hem_unit * panels

    # OP（方向：W/横/X→幅、その他→高さ）
    name_col = pick_col(df_op, ["OP名称"]) or "OP名称"
    price_col= pick_col(df_op, ["金額","単価"]) or "金額"
    dir_col  = pick_col(df_op, ["方向"]) or "方向"

    op_total = 0
    note_ops = []
    for row in (picked_ops_rows or []):
        nm = normalize_string(row.get(name_col, ""))
        if not nm:
            continue
        unit = int(parse_float(row.get(price_col)) or 0)
        dire = normalize_string(row.get(dir_col, "")).upper()
        if unit <= 0:
            continue
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
# サマリ & Excel
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
    ws.cell(r,1,"お見積書（明細）").font = Font(size=14, bold=True); r+=2
    ws.cell(r,1,"見積番号"); ws.cell(r,2, header.get("estimate_no",""))
    ws.cell(r,4,"作成日");   ws.cell(r,5, header.get("date","")); r+=1
    ws.cell(r,1,"得意先"); ws.cell(r,2, header.get("customer_name",""))
    ws.cell(r,4,"部署・支店"); ws.cell(r,5, header.get("branch_name","")); r+=1
    ws.cell(r,1,"事業所"); ws.cell(r,2, header.get("office_name",""))
    ws.cell(r,4,"ご担当"); ws.cell(r,5, header.get("person_name","")); r+=2
    hdr = ["品名","数量","単位","単価","小計","種別","備考"]
    for c,t in enumerate(hdr, start=1):
        cell = ws.cell(r,c,t)
        cell.font = Font(bold=True); cell.fill = head_fill
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
            cell=ws.cell(r,c,v); cell.border=border
            if c in (2,4,5) and isinstance(v,(int,float)):
                cell.number_format = int_fmt if c==2 else yen_fmt
            if c==2: cell.alignment=Alignment(horizontal="right")
            if c in (4,5): cell.alignment=Alignment(horizontal="right")
        if isinstance(it.get("小計"), (int,float)):
            total += it["小計"]
        r+=1
    r+=1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.cell(r,1,"合計").font=Font(bold=True)
    ws.cell(r,5,total).number_format=yen_fmt; ws.cell(r,5).font=Font(bold=True)
    for c in (1,5): ws.cell(r,c).border=border
    bio=BytesIO(); wb.save(bio); bio.seek(0); return bio

# =========================
# UI（Streamlitがある場合のみ）
# =========================
if HAS_STREAMLIT:
    st.set_page_config(page_title="S・MAC 見積", layout="wide")
    st.title("S・MAC 見積アプリ（定型文なし）")

    M = load_all()
    # 正規化
    for k in ["df_gen","df_op","df_catalog","df_perf","df_cus","df_branch","df_office","df_genprice"]:
        df = M[k]
        df.columns = [normalize_string(c) for c in df.columns]

    # 参照変数
    df_gen      = M["df_gen"].copy()
    df_op       = M["df_op"].copy()
    df_catalog  = M["df_catalog"].copy()
    df_perf     = M["df_perf"].copy()
    df_cus      = M["df_cus"].copy()
    df_branch   = M["df_branch"].copy()
    df_office   = M["df_office"].copy()
    df_genprice = M["df_genprice"].copy()

    # 原反の単価結合（製品名で左結合）
    name_g = pick_col(df_gen, ["製品名","品名","名称"]) or df_gen.columns[0]
    wcol   = pick_col(df_gen, ["原反幅(mm)","原反幅","幅","巾"]) or df_gen.columns[1]
    thcol  = pick_col(df_gen, ["厚み","厚さ","t"]) or name_g

    if "製品名" not in df_genprice.columns:
        err_stop("原反単価表.csv の『製品名』列が見つかりません。")

    _df = pd.merge(df_gen, df_genprice[["製品名","単価"]], on="製品名", how="left")
    if _df["単価"].isna().any():
        miss = _df[_df["単価"].isna()]["製品名"].unique().tolist()
        err_stop("原反単価表.csv に単価未登録の製品があります: " + ", ".join(map(str, miss[:10])) + (" ..." if len(miss)>10 else ""))

    df_gen_merged = _df

    # ---- 見積番号（UC37MM-####）
    def gen_estimate_no():
        today = datetime.today(); prefix = f"UC37{today.month:02d}-"
        if "_est_seq" not in st.session_state:
            st.session_state["_est_seq"] = 0
        st.session_state["_est_seq"] += 1
        return prefix + f"{st.session_state['_est_seq']%10000:04d}"

    with st.container():
        st.markdown("---")
        st.markdown("### 得意先情報入力")

        if "estimate_no" not in st.session_state:
            st.session_state["estimate_no"] = gen_estimate_no()
        if "created_disp" not in st.session_state:
            st.session_state["created_disp"] = datetime.today().strftime("%Y/%m/%d")

        # 連動プルダウン（得意先マスタが無ければ手入力に自動フォールバック）
        if not df_cus.empty and not df_branch.empty and not df_office.empty:
            c_cus = pick_col(df_cus, ["社名"]) or df_cus.columns[0]
            c_b_cus = pick_col(df_branch, ["社名"]) or df_branch.columns[0]
            c_b = pick_col(df_branch, ["支店名","支店"]) or df_branch.columns[1]
            c_o_cus = pick_col(df_office, ["社名"]) or df_office.columns[0]
            c_o_b = pick_col(df_office, ["支店名","支店"]) or df_office.columns[1]
            c_o = pick_col(df_office, ["営業所名","営業所"]) or df_office.columns[2]

            cus_list = df_cus[c_cus].dropna().unique().tolist()

            col1, col2, col3 = st.columns([1.0, 1.0, 1.0])
            with col1:
                customer = st.selectbox("社名（得意先名）", cus_list, index=0 if cus_list else None, key="customer")
            with col2:
                branches = df_branch[df_branch[c_b_cus]==customer][c_b].dropna().unique().tolist()
                branch = st.selectbox("支店名", branches, index=0 if branches else None, key="branch")
            with col3:
                offices = df_office[(df_office[c_o_cus]==customer) & (df_office[c_o_b]==branch)][c_o].dropna().unique().tolist()
                office = st.selectbox("営業所名", offices, index=0 if offices else None, key="office")
        else:
            col1, col2, col3 = st.columns([1.0, 1.0, 1.0])
            with col1:
                customer = st.text_input("社名（得意先名）", value=st.session_state.get("customer",""), key="customer")
            with col2:
                branch = st.text_input("支店名", value=st.session_state.get("branch",""), key="branch")
            with col3:
                office = st.text_input("営業所名", value=st.session_state.get("office",""), key="office")

        col4, col5, col6 = st.columns([0.8, 0.6, 1.4])
        with col4:
            est = st.text_input("見積番号", value=st.session_state["estimate_no"])
            st.session_state["estimate_no"] = est
        with col5:
            created = st.text_input("作成日（YYYY/MM/DD）", value=st.session_state["created_disp"])
            st.session_state["created_disp"] = created
        with col6:
            project = st.text_input("物件名", value=st.session_state.get("project_name",""))
            st.session_state["project_name"] = project

        st.caption(
            f"見積番号：{st.session_state['estimate_no']} / 作成日：{st.session_state['created_disp']} / "
            f"得意先：{customer or '（未入力）'} / 支店：{branch or '（未入力）'} / 営業所：{office or '（未入力）'} / 物件名：{project or '（未入力）'}"
        )
        st.markdown("---")

    # 間口
    if "openings" not in st.session_state:
        st.session_state.openings = [1]

    overall_items = []
    overall_total = 0

    def add_total(v):
        globals()["overall_total"] = globals().get("overall_total", 0) + int(v or 0)

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

        # 性能（任意）
        perf_opts = []
        if not M["df_perf"].empty:
            perf_mid_col  = pick_col(M["df_perf"], ["中分類"]) or M["df_perf"].columns[0]
            perf_perf_col = pick_col(M["df_perf"], ["性能"]) or M["df_perf"].columns[1]
            perf_opts = M["df_perf"][M["df_perf"][perf_mid_col]==mid][perf_perf_col].dropna().unique().tolist() if mid else []
        _ = st.selectbox("カーテン性能（任意）", [""]+perf_opts, key=pref+"perf")

        open_mtd = st.radio("片引き/引分け", ["片引き","引分け"], key=pref+"open", horizontal=True)

        # OP
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

        # 計算
        if W>0 and H>0 and CNT>0 and mid:
            r = smac_estimate(mid, open_mtd, W, H, CNT, op_rows)
            if r["ok"]:
                note = f"W{W}×H{H}mm"
                if r["note_ops"]:
                    note += "／OP:" + "・".join(r["note_ops"])
                overall_items.append({
                    "品名": f"S・MACカーテン {mid} {open_mtd}",
                    "数量": CNT, "単位":"式",
                    "単価": r["sell_one"], "小計": r["sell_total"],
                    "種別": "S・MAC",
                    "備考": (f"符号:{mark}／" if mark else "") + note,
                })
                add_total(r["sell_total"])

                # 原価明細
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

    # レイアウト
    st.markdown("---")
    col_a, col_b = st.columns([0.85,0.15])
    with col_b:
        if st.button("＋ 間口を追加", key="add_opening_btn"):
            st.session_state.openings.append(len(st.session_state.openings)+1)
            st.rerun()
    with col_a:
        for i,_ in enumerate(st.session_state.openings, start=1):
            st.markdown(f"#### 間口 {i}")
            render_opening(i)

    # サマリ
    st.markdown("---")
    st.markdown("### 見積サマリ")
    if 'overall_items' in globals() and globals().get('overall_items'):
        cols = ["品名","数量","単位","単価","小計","種別","備考"]
        rows = []
        prev_mark = None
        started = False
        for it in overall_items:
            mark = extract_mark(it.get("備考","")); head = is_opening_head(it)
            if started and ((mark and mark!=prev_mark) or head):
                rows.append({c: "" for c in cols})
            prev_mark = mark if mark else prev_mark; started=True
            rows.append({c: it.get(c, "") for c in cols})
        df_sum = pd.DataFrame(rows, columns=cols)
        def _n(v):
            if v in (None, "", "-"): return ""
            try: return f"{int(v):,}"
            except: return v
        for c in ["数量","単価","小計"]:
            df_sum[c] = df_sum[c].map(_n)
        st.dataframe(df_sum, use_container_width=True, hide_index=True)
    else:
        st.info("明細がありません。")

    # 合計
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
            "person_name":   st.session_state.get("pic",""),  # 任意：担当者名欄が必要ならUI追加してください
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
# CLIテスト（Streamlitなし環境）
# =========================
else:
    def _fixture_frames():
        # テスト専用フィクスチャ（アプリ本体では未使用）
        df_gen = pd.DataFrame({
            "製品名": ["標準0.25t", "標準0.4t"],
            "原反幅(mm)": [1370, 1500],
            "厚み": ["0.25t", "0.4t"],
        })
        df_genprice = pd.DataFrame({
            "製品名": ["標準0.25t", "標準0.4t"],
            "単価": [1800, 2200],
        })
        df_op = pd.DataFrame({
            "OP名称": ["透明窓追加", "補強テープ", "H方向OP"],
            "金額": [1000, 800, 1000],
            "方向": ["W", "W", "H"],
        })
        return df_gen, df_genprice, df_op

    def _prepare_globals_for_tests():
        global df_gen_merged, df_op, name_g, wcol, thcol
        df_gen, df_genprice, df_op_local = _fixture_frames()
        df_gen_merged = pd.merge(df_gen, df_genprice[["製品名","単価"]], on="製品名", how="left")
        df_op = df_op_local
        name_g = "製品名"; wcol = "原反幅(mm)"; thcol = "厚み"

    def run_tests():
        _prepare_globals_for_tests()

        # --- Test 1: 基本計算（片引き） ---
        r1 = smac_estimate("標準0.25t", "片引き", W=2000, H=2000, cnt=1, picked_ops_rows=[])
        assert r1["ok"] is True, "基本計算が失敗"
        assert r1["sell_total"] > 0, "売価が0以下"
        for k in ["原反使用量(m)", "原反幅(mm)", "原反単価(円/m)"]:
            assert k in r1["breakdown"], f"breakdownに{ k }が無い"

        # --- Test 2: 引分けは片引き以上のコストになる傾向 ---
        r2 = smac_estimate("標準0.25t", "引分け", W=2000, H=2000, cnt=1, picked_ops_rows=[])
        assert r2["ok"] is True
        assert r2["sell_total"] >= r1["sell_total"], "引分けが片引きより安くなっている"

        # --- Test 3: OP方向の違い（H方向OPのほうが高くなるケース） ---
        # W < H なので H基準OPのほうが金額が大きいはず
        rW = smac_estimate("標準0.4t", "片引き", W=1500, H=2500, cnt=1,
                           picked_ops_rows=[{"OP名称":"透明窓追加","金額":1000,"方向":"W"}])
        rH = smac_estimate("標準0.4t", "片引き", W=1500, H=2500, cnt=1,
                           picked_ops_rows=[{"OP名称":"H方向OP","金額":1000,"方向":"H"}])
        assert rW["ok"] and rH["ok"], "OPテストの計算失敗"
        assert rH["sell_total"] >= rW["sell_total"], "H方向OPの方が高くならない"

        # --- Test 4: 不正入力 ---
        r_bad = smac_estimate("標準0.25t", "片引き", W=0, H=1000, cnt=1, picked_ops_rows=[])
        assert r_bad["ok"] is False, "不正入力でokになるのはおかしい"

        # --- Test 5: 端数処理（100円切上げ） ---
        assert ceil100(1) == 100 and ceil100(100) == 100 and ceil100(101) == 200, "ceil100 が100円切上げになっていない"
        assert r1["sell_one"] % 100 == 0, "販売単価が100円単位になっていない"

        print("All tests passed.")

    if __name__ == "__main__":
        try:
            run_tests()
        except AssertionError as e:
            print("TEST FAILED:", e)
        except Exception as e:
            print("ERROR:", e)
