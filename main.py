# -*- coding: utf-8 -*-
# === Clean main.py (エラーの元になっている全角記号/壊れた文字列を一掃した最小安定版) ===
from __future__ import annotations
import os, secrets
from datetime import date, datetime
import streamlit as st

# ---------------- 基本UI設定 ----------------
st.set_page_config(page_title="お見積書作成", layout="wide")
st.write("")  # 初期レンダリング安定化

# ---------------- 定型文（安全な ASCII クォートのみ） ----------------
PHRASES = [
    '金具：スチール',
    '金具：ステンレス',
    '戸先側：取手付間仕切りポール',
    '戸尻側：フラットバー固定',
    '戸尻側：取手付間仕切りポール',
    '戸尻側：吊下げ固定',
    '※ シートは収縮を考慮し長めでの出荷となります。現場で裾カット調整してください。',
    '※ カーテン下端は長めでの出荷となります。現場で裾カット調整してください。',
    '※ 下地別途。取付用のビス等は別途。',
    '※ カーテンと間仕切りポール・中間ポールは組込済みです。取手・落し・マグネットは現場で取り付けてください。',
    '※ シートは分割。現場での接続プレートジョイントとなります。',
    '※ 不燃シートは白濁が発生する場合があります。特性上ご了承ください。',
    '※ 一次電源・配線・モール類は別途。施工時現場でご用意ください。',
]

# ---------------- 見積番号 ----------------
FY_BASE_TERM = 37
def _term(d: date) -> int:
    return FY_BASE_TERM + (d.year - 2024) - (0 if d.month >= 10 else 1)

def gen_estimate_no(user_code: str, today: date, used: set[str]) -> str:
    uc = (user_code or "UC").upper()
    mm = f"{today.month:02d}"
    t  = _term(today)
    while True:
        sfx = f"{secrets.randbelow(10000):04d}"
        s = f"{uc}{t}{mm}-{sfx}"
        if s not in used:
            used.add(s); return s

# ---------------- Session 初期化 ----------------
if "used_serials" not in st.session_state: st.session_state["used_serials"] = set()
if "user_code"    not in st.session_state: st.session_state["user_code"] = "UC"
if "estimate_no"  not in st.session_state:
    st.session_state["estimate_no"] = gen_estimate_no(st.session_state["user_code"], date.today(), st.session_state["used_serials"])
if "openings_n"   not in st.session_state: st.session_state["openings_n"] = 1

def regen_estno():
    st.session_state["estimate_no"] = gen_estimate_no(st.session_state.get("user_code","UC"), date.today(), st.session_state["used_serials"])

# ---------------- ヘッダ ----------------
st.markdown("### お見積書作成")
a,b,c,d = st.columns([1,1.2,1,1])
with a: st.text_input("担当者コード", key="user_code")
with b:
    st.text_input("見積番号", key="estimate_no")
with c:
    st.button("見積番号を再生成", on_click=regen_estno, key="regen_est_btn")
with d:
    st.text_input("作成日", value=datetime.today().strftime("%Y/%m/%d"), key="created_disp", disabled=True)

x1,x2 = st.columns([2,1])
with x1: st.text_input("得意先名", key="client")
with x2: st.text_input("担当者名", key="pic")
st.text_input("物件名", key="pj")
st.divider()

# ---------------- 間口 UI ----------------
def render_opening(i: int):
    pref = f"o{i}_"
    st.markdown(f"#### 間口 {i}")

    c1,c2,c3 = st.columns(3)
    with c1: W = st.number_input(f"W(mm)_{i}", min_value=0, step=50, key=pref+"w")
    with c2: H = st.number_input(f"H(mm)_{i}", min_value=0, step=50, key=pref+"h")
    with c3: Q = st.number_input(f"数量_{i}",  min_value=0, step=1,  key=pref+"qty")

    # 大分類（ラジオ）
    large = st.radio(f"種別_{i}", ["","S・MACカーテン","エア・セーブ"], key=pref+"large", horizontal=True)

    # 開閉（ラジオ）
    open_m = st.radio(f"開閉_{i}", ["","片引き","両引き"], key=pref+"open", horizontal=True)

    # 仕様（必要最低限）
    if large == "S・MACカーテン":
        mid = st.selectbox(f"S・MAC 中分類_{i}", ["","0.3t透明","0.5t静電・防炎 片引き"], key=pref+"mid")
        perf = st.selectbox(f"カーテン性能_{i}", ["","透明","防炎","静電"], key=pref+"perf")
        st.session_state[pref+"product"] = mid or "S・MACカーテン"
    elif large == "エア・セーブ":
        airtype = st.radio(f"型式_{i}", ["","MA","MB","MC","ME"], key=pref+"airtype", horizontal=True)
        name = st.radio(f"品名_{i}", ["","標準タイプ","防炎タイプ"], key=pref+"airname", horizontal=True)
        perf = st.selectbox(f"カーテン性能_{i}", ["","透明","防炎"], key=pref+"perf2")
        st.session_state[pref+"product"] = name or airtype or "エア・セーブ"
    else:
        st.session_state[pref+"product"] = ""

    # 部材（横並びラジオ・簡易）
    show_parts = (large == "S・MACカーテン") or (large == "エア・セーブ" and (st.session_state.get(pref+"airtype") or "").startswith("MA"))
    if show_parts:
        st.markdown("##### 部材入力")
        parts = ["", "カーテンレール", "取手付間仕切ポール", "中間ポール", "アルミ押えバー", "落し"]
        rows_key = pref + "parts_rows"
        if rows_key not in st.session_state:
            st.session_state[rows_key] = [{"item":"","qty":1,"unit":""}]
        colA, colB = st.columns([0.9,0.1])
        with colB:
            if st.button("＋", key=pref+"parts_add"):
                st.session_state[rows_key].append({"item":"","qty":1,"unit":""})
                st.rerun()
        with colA: st.caption("各行：品名（ラジオ）／数量／単価（任意）")

        new_rows = []
        for j,row in enumerate(st.session_state[rows_key]):
            r1,r2,r3,r4 = st.columns([1.4,0.45,0.7,0.3])
            with r1:
                item = st.radio(f"品名_{i}-{j+1}", parts, index=parts.index(row.get("item","")) if row.get("item","") in parts else 0,
                                key=f"{pref}p_item_{j}", horizontal=True)
            with r2:
                qty  = st.number_input("数量", min_value=1, value=int(row.get("qty",1)), step=1, key=f"{pref}p_qty_{j}")
            with r3:
                unit = st.number_input("単価（任意）", min_value=0, value=int(row.get("unit",0) or 0), step=100, key=f"{pref}p_unit_{j}")
            with r4:
                delete = st.button("×", key=f"{pref}p_del_{j}")
            if not delete:
                new_rows.append({"item":item,"qty":qty,"unit":unit})
        st.session_state[rows_key] = new_rows
    else:
        st.caption("※ この構成では部材は不要です。")

    # 定型文（部材の下に表示）
    st.markdown("##### 見積書本文（定型句）")
    left, right = st.columns(2)
    sel_texts = []
    half = (len(PHRASES)+1)//2
    with left:
        for idx,p in enumerate(PHRASES[:half]):
            if st.checkbox(p, key=f"{pref}phrL{idx}"): sel_texts.append(p)
    with right:
        for idx,p in enumerate(PHRASES[half:]):
            if st.checkbox(p, key=f"{pref}phrR{idx}"): sel_texts.append(p)
    st.session_state[pref+"phrases"] = sel_texts

def collect_openings():
    items = []
    n = st.session_state.get("openings_n",1)
    for i in range(1, n+1):
        pref = f"o{i}_"
        W = int(st.session_state.get(pref+"w") or 0)
        H = int(st.session_state.get(pref+"h") or 0)
        Q = int(st.session_state.get(pref+"qty") or 0)
        large = st.session_state.get(pref+"large") or ""
        prod  = st.session_state.get(pref+"product") or ""
        items.append({
            "sign":"", "curtain_subtype":large, "open_method":st.session_state.get(pref+"open") or "",
            "product_name":prod, "opening_w":W, "opening_h":H,
            "curtain_w": (W if large=="S・MACカーテン" else 0),
            "curtain_h": (H if large=="S・MACカーテン" else 0),
            "qty": Q if Q>0 else None, "unit": ("式" if Q>0 else None), "unit_price": None,
            "phrases": st.session_state.get(pref+"phrases", [])
        })
    return items

# 追加/削除
b1,b2 = st.columns(2)
with b1:
    if st.button("＋ 間口を追加", key="add_open"): st.session_state["openings_n"] += 1; st.rerun()
with b2:
    if st.button("－ 最後の間口を削除", key="del_open") and st.session_state["openings_n"]>1:
        st.session_state["openings_n"] -= 1; st.rerun()

# 描画
for i in range(1, st.session_state["openings_n"]+1):
    with st.expander(f"間口 {i}", expanded=True):
        render_opening(i)

# ---------------- Excel 保存 ----------------
def _header_from_state():
    created = st.session_state.get("created_disp") or datetime.today().strftime("%Y/%m/%d")
    return {
        "estimate_no": st.session_state.get("estimate_no",""),
        "date":        created.replace("/","-"),
        "customer_name": st.session_state.get("client",""),
        "branch_name":   "",
        "office_name":   "",
        "person_name":   st.session_state.get("pic",""),
    }

def save_excel():
    header = _header_from_state()
    openings = collect_openings()
    if not openings:
        st.error("明細がありません。"); return
    base = os.path.dirname(os.path.abspath(__file__))
    tpl  = os.path.join(base, "見積書テンプレ.xlsx")
    if not os.path.exists(tpl):
        st.error(f"テンプレートが見つかりません: {tpl}"); return
    outdir = os.path.join(base, "data"); os.makedirs(outdir, exist_ok=True)
    fn = f"{datetime.now().strftime('%m%d')}{st.session_state.get('pj') or '見積'}.xlsx"
    out = os.path.join(outdir, fn)
    try:
        from excel_export_com import write_estimate_to_template
    except Exception:
        st.error("excel_export_com.py が見つかりません。"); return
    try:
        write_estimate_to_template(template_path=tpl, output_path=out, header=header, openings=openings)
        st.success(f"Excelを保存しました：{out}")
        with open(out,"rb") as f:
            st.download_button("ダウンロード", f.read(), file_name=os.path.basename(out),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="excel_dl_btn_v1")
    except Exception as e:
        st.error("Excel出力でエラーが発生しました。"); st.exception(e)

st.subheader("Excel保存")
if st.button("Excel保存", key="excel_save_btn_v1"):
    save_excel()
