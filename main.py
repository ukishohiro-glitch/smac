# -*- coding: utf-8 -*-
# main.py — ①大分類ラジオ ②片引き/引分けラジオ ③部材品名ラジオ ④定型文チェック ⑤Excel出力(お見積書(明細)) 対応版
from __future__ import annotations
import os, os.path as osp, secrets, math, re, unicodedata
from datetime import datetime, date
import pandas as pd
import streamlit as st


# ===== 見積サマリ（間口ごとに1行空けて表示） =====
import re as _re_for_summary
import pandas as _pd_for_summary

def _extract_mark(note: str) -> str | None:
    """備考から '符号:XX' を抽出（無ければ None）"""
    if not note:
        return None
    m = _re_for_summary.search(r"符号[:：]\s*([^／\s]+)", str(note))
    return m.group(1) if m else None

def _is_opening_head(it: dict) -> bool:
    """S・MACやエア・セーブの“カーテン”行を間口の先頭候補にする"""
    name = str(it.get("品名", ""))
    kind = str(it.get("種別", ""))
    return ("S・MAC" in kind or "S・MAC" in name) or name.startswith("エア・セーブ")

def render_summary_table(overall_items: list[dict]):
    """overall_items をそのまま使い、間口切り替わりで空行を挿入して表示"""
    cols = ["品名","数量","単位","単価","小計","種別","備考"]
    rows = []
    prev_mark = None
    started = False

    for it in overall_items:
        mark = _extract_mark(it.get("備考", ""))
        is_head = _is_opening_head(it)

        # 一度出力済みで、符号が変わる or “見出し行”に切り替わる → 空行
        if started and ((mark and mark != prev_mark) or is_head):
            rows.append({c: "" for c in cols})

        prev_mark = mark if mark else prev_mark
        started = True

        rows.append({
            "品名": it.get("品名",""),
            "数量": it.get("数量",""),
            "単位": it.get("単位",""),
            "単価": it.get("単価",""),
            "小計": it.get("小計",""),
            "種別": it.get("種別",""),
            "備考": it.get("備考",""),
        })

    df = _pd_for_summary.DataFrame(rows, columns=cols)

    def _fmt_int(v):
        if v in (None, "", "-"):
            return ""
        try:
            return f"{int(v):,}"
        except Exception:
            return str(v)

    for c in ["数量", "単価", "小計"]:
        df[c] = df[c].map(_fmt_int)

    st.markdown("### 見積サマリ")
    st.dataframe(df, use_container_width=True, hide_index=True)

# ===== ユーティリティ =====
def normalize_string(s):
    if s is None: return ""
    if isinstance(s, str):
        s = unicodedata.normalize("NFKC", s)
        s = s.replace("\u3000", " ")
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
    if isinstance(x, (int, float)): return float(x)
    m = NUM_RE.search(str(x))
    return float(m.group()) if m else None

def extract_thickness(s):
    if s is None: return None
    m = re.search(r"(\d+(?:\.\d+)?)\s*(?:t|mm)", str(s).lower())
    return float(m.group(1)) if m else None

def ceil100(x):
    import math
    return int(math.ceil(float(x) / 100.0) * 100)

# ===== データロード（ダミー含む。あなたの環境に合わせて適宜差し替え） =====
@st.cache_data
def load_master():
    # あなたの master.xlsx 読み込みロジックで置換可能
    return {
        "df_gen": pd.DataFrame(),
        "df_op": pd.DataFrame(),
        "df_curtain": pd.DataFrame(),
        "df_perf": pd.DataFrame(),
        "df_ma": pd.DataFrame(),
        "df_mb_tbl": pd.DataFrame(),
        "df_mc": pd.DataFrame(),
        "df_me_curt": pd.DataFrame(),
        "df_me_motor": pd.DataFrame(),
        "df_parts": {},  # dict[str, DataFrame]
    }

m = load_master()
df_gen = m["df_gen"]; df_op = m["df_op"]; df_curtain = m["df_curtain"]; df_perf = m["df_perf"]
df_ma = m["df_ma"]; df_mb_tbl = m["df_mb_tbl"]; df_mc = m["df_mc"]; df_me_curt = m["df_me_curt"]; df_me_motor = m["df_me_motor"]
df_parts = m["df_parts"]

# ===== 画面の共通UI =====
st.set_page_config(page_title="SMAC 見積", layout="wide")
st.title("SMAC 見積アプリ")

def sec_title(t): st.markdown(f"### {t}")

# 合計サマリ
overall_total = 0
overall_items = []
def overall_total_update(v): 
    global overall_total
    overall_total += int(v or 0)

# ===== 間口 =====
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
        large = st.radio("カーテン大分類", ["","S・MACカーテン","エア・セーブ"], key=pref+"large", horizontal=True)

    air_type = None
    middle = small = perf = ""
    rib_note = ""

    with b2:
        if large and large != "エア・セーブ":
            middle = st.selectbox("カーテン中分類", ["基本"], key=pref+"mid")

    with b3:
        if large == "エア・セーブ":
            air_label = st.radio("型式（MA・MB・MC・ME）", ["","MA型折りたたみ式","MB型固定式","MC型スライド式","ME型電動式"], key=pref+"airtype", horizontal=True)
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
            items = ["標準カーテン"]
            if air_type=="MA":
                air_item = st.selectbox("エア・セーブ品名", [""]+items, key=pref+"ma_item")
            elif air_type=="MB":
                air_item = st.selectbox("エア・セーブ品名", [""]+items, key=pref+"mb_item")
            elif air_type=="MC":
                air_item = st.selectbox("エア・セーブ品名", [""]+items, key=pref+"mc_item")
            elif air_type=="ME":
                air_item = st.selectbox("エア・セーブ品名（カーテン）", [""]+items, key=pref+"me_curt")

    with c4:
        if large=="エア・セーブ":
            perf = st.selectbox("カーテン性能", ["","標準"], key=pref+"perf")
        elif large and large!="S・MACカーテン":
            perf = st.selectbox("カーテン性能", ["","標準"], key=pref+"perf2")

    # S・MAC OP（任意・ダミー）
    picked_ops = []
    if large=="S・MACカーテン":
        st.caption("S・MAC OP（任意／金額はカーテンに内包）")
        names = ["透明窓追加","補強テープ"]
        cols = st.columns(3)
        for i, nm in enumerate(names):
            with cols[i%3]:
                if st.checkbox(nm, key=pref+f"smac_op_{i}"):
                    picked_ops.append({"OP名称": nm, "金額": 1000, "方向": "W"})

    # 計算 → サマリ（ここはダミー計算）
    if W>0 and H>0 and CNT>0:
        if large=="S・MACカーテン":
            sm = {"ok": True, "sell_one": 12000, "sell_total": 12000*CNT,
                  "note_ops": [o["OP名称"] for o in picked_ops],
                  "breakdown": {
                      "原反使用量(m)": 3.2, "原反幅(mm)": 1370, "原反単価(円/m)": 1800,
                      "原反材料(1式)": 5760, "裁断賃(1式)": 3000, "幅繋ぎ(1式)": 600, "四方折り返し(1式)": 1800,
                      "OP加算(1式)": 0, "原価(1式)": 11160, "販売単価(1式)": 12000, "販売金額(数量分)": 12000*CNT, "粗利率": 0.07
                  }}
            note = f"W{W}×H{H}mm"
            if sm["note_ops"]:
                note += "／OP：" + "・".join(sm["note_ops"])
            overall_items.append({
                "品名": "S・MACカーテン"
                        + (f" {middle}" if middle else "")
                        + (f" {st.session_state.get(pref+'open')}" if st.session_state.get(pref+'open') else ""),
                "数量": CNT,
                "単位": "式",
                "単価": sm["sell_one"],
                "小計": sm["sell_total"],
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
                        if k not in bd: 
                            continue
                        v = bd[k]
                        if k == "原反使用量(m)":
                            rows.append([k, f"{float(v):.2f} m"])
                        elif k == "原反幅(mm)":
                            rows.append([k, f"{int(v)} mm"])
                        elif k == "原反単価(円/m)":
                            rows.append([k, f"¥{int(v):,}/m"])
                        elif k == "粗利率":
                            rows.append([k, f"{float(v)*100:.1f}%"])
                        else:
                            rows.append([k, f"¥{int(v):,}"])
                    st.dataframe(pd.DataFrame(rows, columns=["項目","金額"]), use_container_width=True, hide_index=True)

        elif large=="エア・セーブ" and air_type:
            price = 10000
            if air_type=="MC":
                rail = 2000
                overall_items.append({
                    "品名": f"エア・セーブ MC型スライド式 標準カーテン",
                    "数量": CNT, "単位":"式", "単価": price, "小計": price*CNT,
                    "種別":"エア・セーブMC",
                    "備考": (f"符号:{mark}／" if mark else "") + f"W{W}×H{H}mm"
                             + (f"／{st.session_state.get(pref+'rib')}" if st.session_state.get(pref+'rib') else "")
                             + (f"／{st.session_state.get(pref+'open')}" if st.session_state.get(pref+'open') else "")
                })
                overall_items.append({
                    "品名": "スライドレール", "数量": 1, "単位":"式", "単価": rail, "小計": rail,
                    "種別":"エア・セーブMC", "備考": "W×2を2000mm刻み"
                })
                overall_total_update(price*CNT + rail)
            else:
                overall_items.append({
                    "品名": f"エア・セーブ {air_type}型 標準カーテン",
                    "数量": CNT, "単位":"式", "単価": price, "小計": price*CNT,
                    "種別": f"エア・セーブ{air_type}", "備考": (f"符号:{mark}／" if mark else "") + f"W{W}×H{H}mm"
                })
                overall_total_update(price*CNT)

    # 部材入力（ダミー UI）
    show_parts = (
        st.session_state.get(pref+"large")=="S・MACカーテン" or
        (st.session_state.get(pref+"large")=="エア・セーブ" and (st.session_state.get(pref+"airtype") or "").startswith("MA"))
    )
    if show_parts:
        st.markdown("##### 部材入力")
        for sheet_name, dfp in {"金物": pd.DataFrame()}.items():
            rows_key = pref + f"{sheet_name}_rows"
            if rows_key not in st.session_state:
                st.session_state[rows_key] = [{"item":"", "qty":1}]

            hcol1, hcol2 = st.columns([0.92, 0.08])
            with hcol1: st.caption(f"【{sheet_name}】")
            with hcol2:
                if st.button("＋", key=pref+f"{sheet_name}_add"):
                    st.session_state[rows_key].append({"item":"", "qty":1}); st.rerun()

            names = ["例：吊り金具A", "例：ビスセットB"]
            updated_rows = []
            for j, rowdata in enumerate(st.session_state[rows_key]):
                st.markdown('<div class="row-compact">', unsafe_allow_html=True)
                col1, col2, col3, col4, col5 = st.columns([1.4,0.45,0.7,0.8,0.28])
                with col1:
                    item_opts = [""] + names
                    current = rowdata["item"] if rowdata["item"] in item_opts else ""
                    sel = st.radio(f"品名 {j+1}", item_opts, index=item_opts.index(current), key=pref+f"{sheet_name}_item_{j}_radio")
                with col2:
                    qty = st.number_input("数量", min_value=1, value=int(rowdata["qty"]), step=1, key=pref+f"{sheet_name}_qty_{j}")
                unit, label, used_m = (0,"",0.0)
                if sel:
                    unit, label, used_m = (500, "ラジオ選択", 0.0)
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
    else:
        st.caption("※このカーテン構成では部材の追加は行いません。")

# ===== 画面構成 =====
st.markdown("---")
sec_title("見積サマリ")
if overall_items:
    render_summary_table(overall_items)
else:
    st.info("明細がありません。")

sec_title("見積金額")
st.metric("税抜合計", f"¥{overall_total:,}")

# ===== ⑤ Excel出力：「お見積書（明細）」へ =====
st.markdown("---")
st.markdown("### Excel出力")
if st.button("Excel保存", key="excel_save_btn_v2"):
    header = {
        "estimate_no": st.session_state.get("estimate_no",""),
        "date":        st.session_state.get("created_disp",""),
        "customer_name": st.session_state.get("client",""),
        "branch_name":   st.session_state.get("branch",""),
        "office_name":   st.session_state.get("office",""),
        "person_name":   st.session_state.get("pic",""),
    }
    try:
        # ここではダミーでExcel生成を省略（あなたの build_estimate_workbook を差し替えてOK）
        from io import BytesIO
        from openpyxl import Workbook
        bio = BytesIO(); wb = Workbook(); wb.active.title="お見積書（明細）"; wb.save(bio); bio.seek(0)
        st.success("Excelを作成しました。")
        st.download_button(
            "ダウンロード（お見積書_明細.xlsx）",
            data=bio.getvalue(),
            file_name="お見積書_明細.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="dl_xlsx_v2",
        )
    except Exception as e:
        st.error("Excel出力でエラーが発生しました。")
        st.exception(e)

# ===== メイン（複数間口のレンダリング） =====
st.markdown("---")
sec_title("間口リスト")
if "openings" not in st.session_state:
    st.session_state.openings = [1]

col_a, col_b = st.columns([0.85, 0.15])
with col_b:
    if st.button("＋ 間口を追加"):
        st.session_state.openings.append(len(st.session_state.openings)+1); st.rerun()

with col_a:
    for i, no in enumerate(st.session_state.openings, start=1):
        st.markdown(f"#### 間口 {i}")
        render_opening(i)
