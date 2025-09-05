# -*- coding: utf-8 -*-
# main.py — 完全版
# 1) カーテン大分類/片引き-引分け/部材の品名 = ラジオボタン
# 2) 定型文 = 部材入力の下にチェックボックス
# 3) S・MAC 原価構成（用語：原反使用量/原反幅/原反単価・裁断賃・幅繋ぎ・四方折り返し）
# 4) 見積サマリ：間口ごとに1行空き（Excel調整不要）
# 5) Excel書き出し：「お見積書（明細）」形式（備考の符号と見出しで自動改行）
#    → “master.xlsx” 等が無くてもダミーデータで動作。後から接続可。

import os, re, math, unicodedata
from io import BytesIO
from datetime import datetime
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# =========================
# 共通ユーティリティ
# =========================
def normalize_string(s):
    if s is None: return ""
    if isinstance(s, str):
        s = unicodedata.normalize("NFKC", s)
        s = s.replace("\u3000", " ")
        s = re.sub(r"\s+", " ", s.strip())
    return s

def pick_col(df: pd.DataFrame, cands):
    if df is None or df.empty: return None
    cols = [normalize_string(c) for c in df.columns]
    df.columns = cols
    for c in cands:
        if c in cols: return c
    # ゆるふわ一致
    for col in cols:
        if any(str(c).lower() in col.lower() for c in cands):
            return col
    return None

NUM_RE = re.compile(r"[-+]?\d[\d,]*\.?\d*")
def parse_float(x):
    if x is None: return None
    if isinstance(x, (int, float)): return float(x)
    m = NUM_RE.search(str(x))
    return float(m.group()) if m else None

def extract_thickness(s):
    if s is None: return None
    m = re.search(r"(\d+(?:\.\d+)?)\s*(?:t|mm)", str(s).lower())
    return float(m.group(1)) if m else None

def ceil100(x):
    return int(math.ceil(float(x) / 100.0) * 100)

# =========================
# マスター読み込み（無ければダミー）
# =========================
@st.cache_data
def load_master():
    """必要なら同ディレクトリに master.xlsx を置いてください。
       シート例: S・MAC原反, S・MAC-OP, エアMA, エアMB, エアMC, エアME-カーテン, エアME-モータ, 部材-金物 …"""
    here = os.getcwd()
    xlsx = os.path.join(here, "master.xlsx")
    data = {}
    if os.path.exists(xlsx):
        try:
            xls = pd.ExcelFile(xlsx)
            def rd(name):
                if name in xls.sheet_names:
                    return pd.read_excel(xlsx, sheet_name=name)
                return pd.DataFrame()
            data = {
                "df_gen": rd("S・MAC原反"),
                "df_op": rd("S・MAC-OP"),
                "df_curtain": rd("S・MAC-カタログ"),
                "df_perf": rd("性能"),
                "df_ma": rd("エアMA"),
                "df_mb_tbl": rd("エアMB"),
                "df_mc": rd("エアMC"),
                "df_me_curt": rd("エアME-カーテン"),
                "df_me_motor": rd("エアME-モータ"),
                "df_parts": {}
            }
            # 部材シートは「部材-」で始まるものを自動収集
            for s in xls.sheet_names:
                if s.startswith("部材-"):
                    data["df_parts"][s.replace("部材-","")] = rd(s)
            return data
        except Exception:
            pass

    # ここからダミー（ファイルが無くても動く）
    df_gen = pd.DataFrame({
        "製品名": ["標準0.25t","標準0.4t"],
        "原反幅(mm)": [1370, 1500],
        "単価": [1800, 2200],
        "厚み": ["0.25t","0.4t"],
    })
    df_op = pd.DataFrame({
        "OP名称": ["透明窓追加","補強テープ"],
        "金額": [1000, 800],
        "方向": ["W","W"],
    })
    df_curtain = pd.DataFrame({"大分類": ["S・MACカーテン"]*2, "中分類":["標準0.25t","標準0.4t"]})
    df_perf = pd.DataFrame({"性能": ["標準"]})

    df_ma = pd.DataFrame({"品名":["標準カーテン"], "性能":["標準"], "基準単価":[9000]})
    df_mb_tbl = pd.DataFrame({"品名":["標準カーテン"], "性能":["標準"], "基準単価":[9500]})
    df_mc = pd.DataFrame({"品名":["標準カーテン"], "性能":["標準"], "基準単価":[10000]})
    df_me_curt = pd.DataFrame({"品名":["標準カーテン"], "性能":["標準"], "基準単価":[10500]})
    df_me_motor = pd.DataFrame({"型式":["モータA","モータB"], "固定価格":[30000, 45000]})
    df_parts = {"金物": pd.DataFrame({"品名":["吊り金具A","ビスセットB"], "単価":[500, 350]})}

    return {
        "df_gen": df_gen, "df_op": df_op, "df_curtain": df_curtain, "df_perf": df_perf,
        "df_ma": df_ma, "df_mb_tbl": df_mb_tbl, "df_mc": df_mc,
        "df_me_curt": df_me_curt, "df_me_motor": df_me_motor,
        "df_parts": df_parts
    }

M = load_master()
df_gen = M["df_gen"]; df_op = M["df_op"]; df_curtain = M["df_curtain"]; df_perf = M["df_perf"]
df_ma = M["df_ma"]; df_mb_tbl = M["df_mb_tbl"]; df_mc = M["df_mc"]
df_me_curt = M["df_me_curt"]; df_me_motor = M["df_me_motor"]
df_parts = M["df_parts"]

# =========================
# S・MAC 計算（用語更新版）
# =========================
def smac_estimate(middle_name: str, open_method: str, W: int, H: int, cnt: int,
                  df_gen: pd.DataFrame, df_op: pd.DataFrame, picked_ops: list[dict]):
    """戻り dict: ok, sell_one, sell_total, note_ops, breakdown"""
    # 単価（必要に応じて調整）
    HEM_UNIT_THIN   = 450     # 四方折り返し（~0.3t）
    HEM_UNIT_THICK  = 550     # 四方折り返し（0.3t~）
    SEAM_UNIT_PER_M = 300     # 幅繋ぎ（1mあたり）
    CUTTING_BASE    = 2000    # 裁断賃（基準）

    res = {"ok": False, "msg": "", "sell_one": 0, "sell_total": 0, "note_ops": [], "breakdown": {}}
    if not middle_name or W<=0 or H<=0 or cnt<=0 or df_gen.empty:
        res["msg"] = "S・MAC：中分類/寸法/数量/原反価格を確認してください。"
        return res

    name_col = pick_col(df_gen, ["製品名","品名","名称"]) or df_gen.columns[0]
    w_col    = pick_col(df_gen, ["原反幅(mm)","原反幅","幅","巾"]) or df_gen.columns[1]
    u_col    = pick_col(df_gen, ["単価","上代","価格","金額"]) or df_gen.columns[2]
    t_col    = pick_col(df_gen, ["厚み","厚さ","t"])

    # 行抽出（完全一致→部分一致）
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

    # 寸法補正
    if open_method == "片引き":
        cur_w = W * 1.05; panels = 1
    else:
        cur_w = (W/2) * 1.05; panels = 2
    cur_h = H + 50

    # 原反使用量（1間口）
    length_per_panel_m = (cur_h * 1.2) / 1000.0
    joints = math.ceil(cur_w / gen_width)                # 1パネル内の継ぎ数量
    raw_len_m = length_per_panel_m * joints * panels     # ★原反使用量[m]
    raw_one = gen_price * raw_len_m                      # 原反材料（1式）

    # 裁断賃（カット工賃→裁断賃）
    cutting_one = (CUTTING_BASE if joints <= 3 else CUTTING_BASE + 1000) * panels

    # 幅繋ぎ（縦継ぎ）
    seams_total = max(0, joints - 1) * panels
    seam_one = math.ceil(cur_h/1000.0) * SEAM_UNIT_PER_M * seams_total

    # 四方折り返し
    hem_unit = HEM_UNIT_THIN if (thick is not None and thick <= 0.3) else HEM_UNIT_THICK
    hem_perimeter_m = (cur_w + cur_w + cur_h + cur_h) / 1000.0
    fourfold_one = math.ceil(hem_perimeter_m) * hem_unit * panels

    # OP（方向 W/横/X → 幅基準、それ以外 → 高さ基準）
    note_ops, op_total = [], 0
    for op in (picked_ops or []):
        name = normalize_string(op.get("OP名称","")); unit = int(parse_float(op.get("金額")) or 0)
        dire = normalize_string(op.get("方向","")).upper()
        if not name or unit<=0: continue
        base_mm = cur_w if dire in ["W","横","X"] else cur_h
        units_1000 = math.ceil(base_mm/1000.0)
        sub = units_1000 * unit * panels * cnt
        op_total += sub; note_ops.append(name)
    op_one = op_total / cnt if cnt else 0

    # 原価→売価
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
# エア・セーブ（簡易）
# =========================
def area_price(df: pd.DataFrame, item: str, perf: str, W: int, H: int, CNT: int):
    if df is None or df.empty:
        base = max(8000, int((W*H/1_000_000)*3500))
        return {"ok": True, "price_one": base, "total": base*CNT}
    name_col = pick_col(df, ["品名","製品名","名称","品番","型式"]) or df.columns[0]
    k_col = pick_col(df, ["基準単価","単価","価格"]) or (df.columns[1] if len(df.columns)>1 else None)
    hit = df[df[name_col]==item]
    if hit.empty and name_col:
        hit = df[df[name_col].astype(str).str.contains(re.escape(item), na=False)]
    if hit.empty:
        base = max(8000, int((W*H/1_000_000)*3500))
    else:
        base = int(parse_float(hit.iloc[0][k_col]) or 9000)
    return {"ok": True, "price_one": base, "total": base*CNT}

def mc_slide_rail_price(W: int, CNT: int):
    # W×2 を 2000mm 刻み → 1本=¥2000 と仮定
    length = W*2
    pieces = math.ceil(length / 2000)
    return pieces * 2000

# =========================
# サマリ（間口ごとに1行空けて表示）
# =========================
def _extract_mark(note: str) -> str | None:
    if not note: return None
    m = re.search(r"符号[:：]\s*([^／\s]+)", str(note))
    return m.group(1) if m else None

def _is_opening_head(it: dict) -> bool:
    name = str(it.get("品名","")); kind = str(it.get("種別",""))
    return ("S・MAC" in kind or "S・MAC" in name) or name.startswith("エア・セーブ")

def render_summary_table(overall_items: list[dict]):
    cols = ["品名","数量","単位","単価","小計","種別","備考"]
    rows = []
    prev_mark = None
    started = False
    for it in overall_items:
        mark = _extract_mark(it.get("備考","")); is_head = _is_opening_head(it)
        if started and ((mark and mark != prev_mark) or is_head):
            rows.append({c: "" for c in cols})  # 空行
        prev_mark = mark if mark else prev_mark; started = True
        rows.append({
            "品名": it.get("品名",""),
            "数量": it.get("数量",""),
            "単位": it.get("単位",""),
            "単価": it.get("単価",""),
            "小計": it.get("小計",""),
            "種別": it.get("種別",""),
            "備考": it.get("備考",""),
        })
    df = pd.DataFrame(rows, columns=cols)
    def _fmt(v):
        if v in (None,"","-"): return ""
        try: return f"{int(v):,}"
        except: return str(v)
    for c in ["数量","単価","小計"]:
        df[c] = df[c].map(_fmt)
    st.markdown("### 見積サマリ")
    st.dataframe(df, use_container_width=True, hide_index=True)

# =========================
# Excel 出力（テンプレ不要の自立版）
# =========================
def build_estimate_workbook(header: dict, items: list[dict]) -> BytesIO:
    wb = Workbook(); ws = wb.active; ws.title = "お見積書（明細）"
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor="F2F2F2")
    yen_fmt = '"¥"#,##0'; int_fmt = '#,##0'
    widths = [36, 8, 7, 11, 12, 12, 56]  # 品名,数量,単位,単価,小計,種別,備考
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
    for it in items:  # 受け取った順で出力
        mark = _extract_mark(it.get("備考","")); is_head = _is_opening_head(it)
        if started and ((mark and mark!=prev_mark) or is_head):
            r += 1  # 間口切替で1行空き
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
# 画面構成
# =========================
st.set_page_config(page_title="SMAC 見積", layout="wide")
st.title("SMAC 見積アプリ")

def sec_title(t): st.markdown(f"### {t}")

overall_items = []
overall_total = 0
def overall_total_update(v):
    nonlocal_total = int(v or 0)
    globals()['overall_total'] = globals().get('overall_total',0) + nonlocal_total

PHRASES = [
    "現地採寸・施工費別途", "納期はご発注後◯週間", "運賃別途", "下地・取付条件は別途確認",
    "色味・柄は現地確認", "夜間作業・休日作業は割増"
]

# ====== 間口1スロット ======
def render_opening(idx: int):
    pref = f"o{idx}_"

    a1,a2,a3,a4 = st.columns([0.7,1.0,1.0,0.7])
    with a1: mark = st.text_input("符号", key=pref+"mark")
    with a2: W = st.number_input("間口W (mm)", min_value=0, value=0, step=50, key=pref+"w")
    with a3: H = st.number_input("間口H (mm)", min_value=0, value=0, step=50, key=pref+"h")
    with a4: CNT = st.number_input("数量", min_value=1, value=1, step=1, key=pref+"cnt")

    sec_title("カーテン入力")

    b1,b2,b3 = st.columns([1.0,1.0,1.0])
    with b1:
        large_list = ["S・MACカーテン","エア・セーブ"]
        large = st.radio("カーテン大分類", [""]+large_list, key=pref+"large", horizontal=True)

    air_type = None
    middle = small = perf = ""
    rib_note = ""

    with b2:
        if large and large != "エア・セーブ":
            # マスターがあればそこから、中分類候補を抽出
            mids = []
            if not df_curtain.empty and "中分類" in df_curtain.columns:
                mids = df_curtain["中分類"].dropna().unique().tolist()
            else:
                mids = [normalize_string(c) for c in (df_gen["製品名"] if "製品名" in df_gen.columns else [])]
                if not mids: mids = ["標準0.25t","標準0.4t"]
            middle = st.selectbox("カーテン中分類", [""]+mids, key=pref+"mid")

    with b3:
        if large == "エア・セーブ":
            air_label = st.radio("型式（MA・MB・MC・ME）",
                                 ["","MA型折りたたみ式","MB型固定式","MC型スライド式","ME型電動式"],
                                 key=pref+"airtype", horizontal=True)
            air_type = air_label[:2] if air_label else None
        else:
            small = st.selectbox("カーテン小分類", [""], key=pref+"small")

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
            # マスターから候補
            src = {"MA":df_ma,"MB":df_mb_tbl,"MC":df_mc,"ME":df_me_curt}.get(air_type, pd.DataFrame())
            name_col = pick_col(src, ["品名","製品名","名称","品番","型式"]) or (src.columns[0] if not src.empty else None)
            items = src[name_col].dropna().unique().tolist() if name_col else ["標準カーテン"]
            air_item = st.selectbox("エア・セーブ品名", [""]+items, key=pref+"air_item")

    with c4:
        if large=="エア・セーブ":
            perf_all = df_perf["性能"].dropna().unique().tolist() if "性能" in df_perf.columns else ["標準"]
            perf = st.selectbox("カーテン性能", [""]+perf_all, key=pref+"perf")
        elif large and large!="S・MACカーテン":
            perf_opts = df_perf["性能"].dropna().unique().tolist() if "性能" in df_perf.columns else ["標準"]
            perf = st.selectbox("カーテン性能", [""]+perf_opts, key=pref+"perf2")

    # S・MAC OP（任意）
    picked_ops = []
    if large=="S・MACカーテン":
        st.caption("S・MAC OP（任意／金額はカーテンに内包）")
        if df_op is not None and not df_op.empty and all(c in df_op.columns for c in ["OP名称","金額","方向"]):
            names = df_op["OP名称"].dropna().unique().tolist()
            cols = st.columns(3)
            for i, nm in enumerate(names):
                with cols[i%3]:
                    if st.checkbox(nm, key=pref+f"smac_op_{i}"):
                        picked_ops.append(df_op[df_op["OP名称"]==nm].iloc[0].to_dict())
        else:
            # ダミー
            names = ["透明窓追加","補強テープ"]
            cols = st.columns(3)
            for i,nm in enumerate(names):
                with cols[i%3]:
                    if st.checkbox(nm, key=pref+f"smac_op_dmy_{i}"):
                        picked_ops.append({"OP名称":nm,"金額":1000,"方向":"W"})

    # ========= 計算 → サマリ =========
    if W>0 and H>0 and CNT>0:
        if large=="S・MACカーテン":
            sm = smac_estimate(middle or "", st.session_state.get(pref+"open") or "片引き",
                               W, H, CNT, df_gen, df_op, picked_ops)
            if sm["ok"]:
                note = f"W{W}×H{H}mm"
                if sm["note_ops"]: note += "／OP：" + "・".join(sm["note_ops"])
                overall_items.append({
                    "品名": "S・MACカーテン"
                            + (f" {middle}" if middle else "")
                            + (f" {st.session_state.get(pref+'open')}" if st.session_state.get(pref+'open') else ""),
                    "数量": CNT, "単位": "式",
                    "単価": sm["sell_one"], "小計": sm["sell_total"],
                    "種別": "S・MAC",
                    "備考": (f"符号:{mark}／" if mark else "") + note
                })
                overall_total_update(sm["sell_total"])

                # 原価構成（1間口あたり）
                bd = sm.get("breakdown", {})
                if bd:
                    with st.expander("原価構成（1間口あたり）", expanded=False):
                        order = [
                            "原反使用量(m)", "原反幅(mm)", "原反単価(円/m)",
                            "原反材料(1式)", "裁断賃(1式)", "幅繋ぎ(1式)", "四方折り返し(1式)", "OP加算(1式)",
                            "原価(1式)", "販売単価(1式)", "販売金額(数量分)", "粗利率"
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
                        st.dataframe(pd.DataFrame(rows, columns=["項目","金額"]),
                                     use_container_width=True, hide_index=True)
            else:
                st.warning(sm.get("msg") or "S・MACの計算に失敗しました。")

        elif large=="エア・セーブ" and air_type:
            if air_type in ["MA","MB","MC"]:
                src = {"MA":df_ma,"MB":df_mb_tbl,"MC":df_mc}[air_type]
                r = area_price(src, air_item or "標準カーテン", perf or "標準", W, H, CNT)
                if r["ok"]:
                    if air_type=="MC":
                        rail = mc_slide_rail_price(W, CNT)
                        overall_items.append({
                            "品名": f"エア・セーブ MC型スライド式 {air_item or '標準カーテン'}",
                            "数量": CNT, "単位":"式", "単価": r["price_one"], "小計": r["price_one"]*CNT,
                            "種別":"エア・セーブMC",
                            "備考": (f"符号:{mark}／" if mark else "") + f"W{W}×H{H}mm"
                                    + (f"／{st.session_state.get(pref+'rib')}" if st.session_state.get(pref+'rib') else "")
                                    + (f"／{st.session_state.get(pref+'open')}" if st.session_state.get(pref+'open') else "")
                        })
                        overall_items.append({
                            "品名": "スライドレール", "数量": 1, "単位":"式",
                            "単価": rail, "小計": rail,
                            "種別":"エア・セーブMC", "備考": "W×2を2000mm刻み"
                        })
                        overall_total_update(r["price_one"]*CNT + rail)
                    else:
                        label = {"MA":"MA型折りたたみ式","MB":"MB型固定式"}[air_type]
                        memo = (f"符号:{mark}／" if mark else "") + f"W{W}×H{H}mm"
                        if air_type=="MB" and st.session_state.get(pref+"rib"):
                            memo += f"／{st.session_state.get(pref+'rib')}"
                        overall_items.append({
                            "品名": f"エア・セーブ {label} {air_item or '標準カーテン'}",
                            "数量": CNT, "単位":"式", "単価": r["price_one"], "小計": r["total"],
                            "種別": f"エア・セーブ{air_type}", "備考": memo
                        })
                        overall_total_update(r["total"])
            elif air_type=="ME":
                total_me=0
                if True:
                    curt_df = df_me_curt if not df_me_curt.empty else df_mc
                    r = area_price(curt_df, air_item or "標準カーテン", perf or "標準", W, H, CNT)
                    if r["ok"]:
                        overall_items.append({
                            "品名": f"エア・セーブ ME型電動式 カーテン {air_item or '標準カーテン'}",
                            "数量": CNT, "単位":"式", "単価": r["price_one"], "小計": r["total"],
                            "種別":"エア・セーブME(カーテン)",
                            "備考": (f"符号:{mark}／" if mark else "") + f"W{W}×H{H}mm"
                        })
                        total_me += r["total"]
                # モータ選択は省略（必要なら追加）
                if total_me>0: overall_total_update(total_me)

    # ===== 部材入力 =====
    show_parts = (
        st.session_state.get(pref+"large")=="S・MACカーテン" or
        (st.session_state.get(pref+"large")=="エア・セーブ" and (st.session_state.get(pref+"airtype") or "").startswith("MA"))
    )
    if show_parts:
        st.markdown("##### 部材入力")
        part_sources = df_parts if df_parts else {"金物": pd.DataFrame({"品名":["吊り金具A","ビスセットB"], "単価":[500,350]})}
        for sheet_name, dfp in part_sources.items():
            rows_key = pref + f"{sheet_name}_rows"
            if rows_key not in st.session_state:
                st.session_state[rows_key] = [{"item":"", "qty":1}]

            hcol1, hcol2 = st.columns([0.92, 0.08])
            with hcol1: st.caption(f"【{sheet_name}】")
            with hcol2:
                if st.button("＋", key=pref+f"{sheet_name}_add"):
                    st.session_state[rows_key].append({"item":"", "qty":1}); st.rerun()

            name_col = pick_col(dfp, ["品名","製品名","名称","品番"]) or (dfp.columns[0] if not dfp.empty else None)
            names = dfp[name_col].dropna().unique().tolist() if name_col else ["吊り金具A","ビスセットB"]
            updated_rows = []
            for j, rowdata in enumerate(st.session_state[rows_key]):
                st.markdown('<div class="row-compact">', unsafe_allow_html=True)
                col1, col2, col3, col4, col5 = st.columns([1.4,0.45,0.7,0.8,0.28])
                with col1:
                    item_opts = [""] + names
                    current = rowdata["item"] if rowdata["item"] in item_opts else ""
                    sel = st.radio(f"品名 {j+1}", item_opts, index=item_opts.index(current), key=pref+f"{sheet_name}_item_{j}_radio", horizontal=True)
                with col2:
                    qty = st.number_input("数量", min_value=1, value=int(rowdata["qty"]), step=1, key=pref+f"{sheet_name}_qty_{j}")
                unit, label, used_m = (0,"",0.0)
                if sel:
                    if not dfp.empty and name_col and sel in dfp[name_col].values:
                        # マスターあり
                        rrow = dfp[dfp[name_col]==sel].iloc[0]
                        unit = int(parse_float(rrow.get("単価")) or 0)
                    else:
                        unit = 500 if "吊" in sel or "金具" in sel else 350
                    label = "部材"
                with col3: st.text_input("単価（自動）", value=(str(unit) if unit else ""), disabled=True, key=pref+f"{sheet_name}_unit_{j}")
                with col4: st.text_input("使用m数（自動）", value="", disabled=True, key=pref+f"{sheet_name}_m_{j}")
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

        # 定型文
        st.markdown("##### 定型文")
        phr_sel=[]
        colL, colR = st.columns(2); half=(len(PHRASES)+1)//2
        with colL:
            for i,p in enumerate(PHRASES[:half]):
                if st.checkbox(p, key=pref+f"phr_L_{i}"): phr_sel.append(p)
        with colR:
            for i,p in enumerate(PHRASES[half:]):
                if st.checkbox(p, key=pref+f"phr_R_{i}"): phr_sel.append(p)
        if phr_sel:
            overall_items.append({
                "品名":"（定型文）","数量":"","単位":"","単価":"","小計":"","種別":"メモ","備考":" / ".join(phr_sel)
            })
    else:
        st.caption("※このカーテン構成では部材の追加は行いません。")

# =========================
# 画面：間口・サマリ・Excel
# =========================
st.markdown("---")
if "openings" not in st.session_state:
    st.session_state.openings = [1]

sec_title("間口リスト")
col_a, col_b = st.columns([0.85,0.15])
with col_b:
    if st.button("＋ 間口を追加", key="add_opening_btn"):
        st.session_state.openings.append(len(st.session_state.openings)+1); st.rerun()
with col_a:
    for i,_ in enumerate(st.session_state.openings, start=1):
        st.markdown(f"#### 間口 {i}")
        render_opening(i)

st.markdown("---")
sec_title("見積サマリ")
if overall_items:
    render_summary_table(overall_items)
else:
    st.info("明細がありません。")

sec_title("見積金額")
st.metric("税抜合計", f"¥{globals().get('overall_total',0):,}")

st.markdown("---")
st.markdown("### Excel出力")
if st.button("Excel保存", key="excel_save_btn_v3"):
    header = {
        "estimate_no": st.session_state.get("estimate_no",""),
        "date":        st.session_state.get("created_disp", datetime.today().strftime("%Y/%m/%d")),
        "customer_name": st.session_state.get("client",""),
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
            key="dl_xlsx_v3"
        )
    except Exception as e:
        st.error("Excel出力でエラーが発生しました。")
        st.exception(e)
