# -*- coding: utf-8 -*-
# main.py — UI層（Streamlit）
# - 得意先情報入力（見積書0: J1/J3/A6/A7/A8/B17）
# - 製品情報入力（master.xlsx連動＋手入力フォールバック）
# - 明細=入力順で「見積書1〜5」12〜44行（11行目は触らない）
# - 1ページ=33行。1間口<=33はページ跨ぎ禁止、>33は跨ぎ可。総ページ>5でエラー
# - 見積書0は21〜44に「間口合計」を連続記載、直後に「運賃・梱包」（空行なし）
# - 梱包必須: S・MAC / エア・セーブMA。MB/MC/MEは任意
# - Excelはテンプレ罫線保持のみ（テンプレ不足ならエラー）

import os
import os.path as osp
import re
import secrets
from pathlib import Path
from datetime import date, datetime
from typing import Dict, Any, List

import streamlit as st
from openpyxl import load_workbook

# __file__ 非対応環境でも動作
try:
    APP_DIR = Path(__file__).parent
except NameError:
    APP_DIR = Path.cwd()

TEMPLATE_BOOK = APP_DIR / "お見積書（明細）.xlsx"
MASTER_BOOK   = APP_DIR / "master.xlsx"

# 業務ロジック / Excel出力
from quote_logic import clean_openings, plan_paging, validate
from excel_export import export_quotation_book_preserve


# --------------------
# 共通ユーティリティ
# --------------------
def _next_estimate_no() -> str:
    used = st.session_state.setdefault("used_estnos", set())
    for _ in range(2000):
        no = "3709-" + f"{secrets.randbelow(100000):05d}"
        if no not in used:
            used.add(no)
            return no
    return "3709-" + datetime.now().strftime("%H%M%S")[-5:]


def _ensure_session():
    ss = st.session_state
    # 得意先情報
    ss.setdefault("estimate_no", _next_estimate_no())
    ss.setdefault("file_title", "見積")
    ss.setdefault("customer_name", "")
    ss.setdefault("branch_name", "")
    ss.setdefault("office_name", "")
    ss.setdefault("rep_name", "")
    ss.setdefault("project_name", "")
    # 製品追加UI（master 連動）
    ss.setdefault("catalog", _load_master(MASTER_BOOK))
    ss.setdefault("cat_sel", "")
    ss.setdefault("search_query", "")
    ss.setdefault("pick_qty", 1)
    ss.setdefault("pick_opening_idx", 1)
    # 手入力フォールバック
    ss.setdefault("free_name", "")
    ss.setdefault("free_qty", 1)
    ss.setdefault("free_unit", "式")
    ss.setdefault("free_price", 0)
    ss.setdefault("free_opening_idx", 1)
    # 明細（間口単位）
    if "openings" not in ss:
        ss.openings = [{
            "name": "開口1",
            "items": [],      # {"品名","数量","単位","単価"}
            "teikeiku": []    # 定型句（文字列）
        }]
    # 運賃・梱包
    ss.setdefault("ship_option", "（路線便時間指定不可）")
    ss.setdefault("ship_price", 0)


def _load_master(path: Path) -> Dict[str, List[Dict[str, Any]]]:
    """master.xlsx をシート単位で読み込み。
       必須列: 品名 / 単位 / 単価（見出しは近似一致も許容）
    """
    cat: Dict[str, List[Dict[str, Any]]] = {}
    if not path.exists():
        return cat
    wb = load_workbook(str(path), data_only=True, read_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        header = [str(x or "").strip() for x in rows[0]]
        idx = {h: i for i, h in enumerate(header)}

        def gi(*names):
            # 完全一致 → 部分一致の順で拾う
            for nm in names:
                if nm in idx:
                    return idx[nm]
            for nm in names:
                for h, i in idx.items():
                    if nm in h:
                        return i
            return None

        c_name = gi("品名", "商品名", "名称", "製品名", "品番")
        if c_name is None:
            continue
        c_unit = gi("単位", "Unit")
        c_price = gi("単価", "基準単価", "上代", "価格", "金額")
        c_model = gi("型番", "型式", "コード")

        items: List[Dict[str, Any]] = []
        for r in rows[1:]:
            if not r:
                continue
            name = str(r[c_name] if c_name < len(r) else "").strip()
            if not name:
                continue
            unit = str(r[c_unit] if (c_unit is not None and c_unit < len(r)) else "式").strip() or "式"
            rawp = r[c_price] if (c_price is not None and c_price < len(r)) else 0
            try:
                price = int(float(str(rawp).replace(",", "")))
                if price < 0:
                    price = 0
            except Exception:
                price = 0
            model = str(r[c_model] if (c_model is not None and c_model < len(r)) else "").strip()
            items.append({"品名": name, "単位": unit, "単価": price, "型番": model})
        if items:
            cat[ws.title] = items
    return cat


# --------------------
# UI: 得意先情報
# --------------------
def _customer_block():
    st.set_page_config(layout="wide", page_title="お見積書作成")
    st.title("お見積書作成")

    c1, c2, c3 = st.columns([1.2, 0.9, 1.0])
    with c1:
        st.text_input("見積番号（固定: 3709-xxxxx）", value=st.session_state.estimate_no,
                      key="estimate_no", disabled=True)
    with c2:
        st.text_input("保存ファイル名（接頭辞）", value=st.session_state.file_title, key="file_title")
    with c3:
        st.text_input("作成日", value=date.today().strftime("%Y/%m/%d"),
                      key="created_disp", disabled=True)

    d1, d2, d3, d4 = st.columns(4)
    with d1:
        st.text_input("得意先名（見積書0:A6）", value=st.session_state.customer_name, key="customer_name")
    with d2:
        st.text_input("支店名（見積書0:A7 左）", value=st.session_state.branch_name, key="branch_name")
    with d3:
        st.text_input("営業所名（見積書0:A7 右）", value=st.session_state.office_name, key="office_name")
    with d4:
        st.text_input("担当者名（半角英数2〜4）", value=st.session_state.rep_name, key="rep_name")
        rep = st.session_state.get("rep_name") or ""
        if rep and not re.fullmatch(r"[A-Za-z0-9]{2,4}", rep):
            st.error("担当者名は半角英数字2〜4文字で入力してください。")

    st.text_input("物件名（見積書0:B17）", value=st.session_state.project_name, key="project_name")
    st.divider()


# --------------------
# UI: 製品情報（master連動）
# --------------------
def _product_master_block():
    st.header("製品情報入力（master.xlsx 連動）")

    cat = st.session_state.catalog
    if not cat:
        st.error("master.xlsx が見つかりません。テンプレ保持の出力要件のため、master連動は必須です。")
        return

    cat_names = sorted(cat.keys())
    if not st.session_state.cat_sel:
        st.session_state.cat_sel = cat_names[0] if cat_names else ""

    cols = st.columns([2, 2, 3, 1, 1])
    with cols[0]:
        csel = st.selectbox("カテゴリ（シート名）", options=cat_names, index=cat_names.index(st.session_state.cat_sel))
        st.session_state.cat_sel = csel
    items = cat.get(csel, [])

    with cols[1]:
        st.text_input("キーワード検索", value=st.session_state.get("search_query", ""), key="search_query")
    q = (st.session_state.get("search_query") or "").strip()
    filtered = [it for it in items if q.lower() in (it["品名"] + " " + (it.get("型番") or "")).lower()] if q else items

    names = [f'{it["品名"]} 〔{it.get("型番","") or "-"} / {it["単価"]}円 / {it["単位"]}〕' for it in filtered] or ["（該当なし）"]
    with cols[2]:
        sel = st.selectbox("部材を選択", options=names, key="part_sel")
    with cols[3]:
        st.number_input("数量", min_value=1, step=1, value=int(st.session_state.pick_qty), key="pick_qty")
    with cols[4]:
        st.number_input("追加先 間口No", min_value=1, step=1, value=int(st.session_state.pick_opening_idx), key="pick_opening_idx")

    def _add_from_master():
        if not filtered:
            return
        i = names.index(sel) if sel in names else -1
        if i < 0:
            return
        item = dict(filtered[i])
        qty = int(st.session_state.pick_qty)
        opi = int(st.session_state.pick_opening_idx)
        while len(st.session_state.openings) < opi:
            st.session_state.openings.append({"name": f"開口{len(st.session_state.openings)+1}", "items": [], "teikeiku": []})
        st.session_state.openings[opi-1]["items"].append(
            {"品名": item["品名"], "数量": qty, "単位": item["単位"], "単価": item["単価"]}
        )

    st.button("＋ この部材を明細へ追加", on_click=_add_from_master, key="btn_add_master")
    st.divider()


# --------------------
# UI: 手入力（フォールバック）
# --------------------
def _product_free_block():
    st.header("製品情報入力（未登録時の手入力フォールバック）")
    cols = st.columns([5, 1, 1, 2, 1])
    with cols[0]:
        st.text_input("品名（A列）", value=st.session_state.free_name, key="free_name")
    with cols[1]:
        st.number_input("数量（F列）", min_value=0, step=1, value=int(st.session_state.free_qty), key="free_qty")
    with cols[2]:
        st.text_input("単位（G列）", value=st.session_state.free_unit, key="free_unit")
    with cols[3]:
        st.number_input("単価（H列・円）", min_value=0, step=100, value=int(st.session_state.free_price), key="free_price")
    with cols[4]:
        st.number_input("間口No", min_value=1, step=1, value=int(st.session_state.free_opening_idx), key="free_opening_idx")

    def _add_free():
        name = (st.session_state.get("free_name") or "").strip()
        if not name:
            st.warning("品名を入力してください。")
            return
        qty = int(st.session_state.get("free_qty") or 0)
        unit = (st.session_state.get("free_unit") or "式").strip() or "式"
        price = int(st.session_state.get("free_price") or 0)
        opi = int(st.session_state.get("free_opening_idx") or 1)
        while len(st.session_state.openings) < opi:
            st.session_state.openings.append({"name": f"開口{len(st.session_state.openings)+1}", "items": [], "teikeiku": []})
        st.session_state.openings[opi-1]["items"].append({"品名": name, "数量": qty, "単位": unit, "単価": price})

    st.button("＋ 手入力の行を明細へ追加", on_click=_add_free, key="btn_add_free")
    st.caption("※ masterに未登録の品はこの欄から追加します。")
    st.divider()


# --------------------
# UI: 明細編集（間口）
# --------------------
def _openings_block():
    st.header("明細編集（間口ごと）")
    st.caption("※ 入力順で出力。11行目は触らず、12〜44行を使用。")
    for oi, op in enumerate(st.session_state.openings):
        with st.expander(f"間口 {oi+1}：{op.get('name') or '(名称未設定)'}", expanded=True):
            op["name"] = st.text_input("間口名", value=op.get("name", f"開口{oi+1}"), key=f"op{oi}_name")

            del_idx = None
            for ri, row in enumerate(op["items"]):
                cols = st.columns([5, 1, 1, 2, 0.9])
                with cols[0]:
                    row["品名"] = st.text_input("品名", value=str(row.get("品名","")), key=f"op{oi}_r{ri}_name")
                with cols[1]:
                    row["数量"] = st.number_input("数量", min_value=0, step=1, value=int(row.get("数量",0)), key=f"op{oi}_r{ri}_qty")
                with cols[2]:
                    row["単位"] = st.text_input("単位", value=str(row.get("単位","式")), key=f"op{oi}_r{ri}_unit")
                with cols[3]:
                    row["単価"] = st.number_input("単価（円）", min_value=0, step=100, value=int(row.get("単価",0)), key=f"op{oi}_r{ri}_price")
                with cols[4]:
                    if st.button("−", key=f"op{oi}_r{ri}_del"):
                        del_idx = ri
                st.write("")
            if del_idx is not None and 0 <= del_idx < len(op["items"]):
                del op["items"][del_idx]

            if st.button("＋ 行を追加（空行）", key=f"op{oi}_addrow"):
                op["items"].append({"品名":"", "数量":0, "単位":"式", "単価":0})

            tei = "\n".join(op.get("teikeiku") or [])
            tei = st.text_area("定型句（1行=1文、間口末尾のA列に出力）", value=tei, key=f"op{oi}_teikei", height=80)
            op["teikeiku"] = [s for s in tei.splitlines() if s.strip()]

            last = st.columns([6, 1])[1]
            with last:
                if st.button("× 間口削除", key=f"op{oi}_del") and len(st.session_state.openings) > 1:
                    st.session_state.openings.pop(oi)
                    st.experimental_rerun()

    if st.button("＋ 間口を追加", key="add_opening"):
        st.session_state.openings.append({"name": f"開口{len(st.session_state.openings)+1}", "items": [], "teikeiku": []})

    st.divider()

    # 正規化して保持 & ページ数概算
    st.session_state.overall_items = clean_openings([
        {"name": op.get("name",""), "items": list(op.get("items") or []), "teikeiku": list(op.get("teikeiku") or [])}
        for op in st.session_state.openings
    ])
    if st.session_state.overall_items:
        try:
            pages = plan_paging(st.session_state.overall_items, rows_per_page=33)
            if pages > 5:
                st.warning(f"この入力だと明細は約 {pages} ページ（上限5）。保存時にエラーになります。")
        except Exception:
            pass


# --------------------
# UI: 運賃・梱包（見積書0のみ）
# --------------------
def _shipping_block():
    st.header("運賃・梱包（見積書0の末尾に記載／明細には出しません）")
    a, b = st.columns([2, 2])
    with a:
        st.radio("配送条件", options=["（路線便時間指定不可）", "（現場搬入時間指定可）"],
                 index=0 if st.session_state.ship_option == "（路線便時間指定不可）" else 1,
                 key="ship_option", horizontal=True)
    with b:
        st.number_input("金額（円）", min_value=0, step=100,
                        value=int(st.session_state.ship_price), key="ship_price")
    st.divider()


# --------------------
# 保存（テンプレ保持のみ）
# --------------------
def _save_block():
    st.header("保存")
    if st.button("Excel保存（テンプレ罫線保持）", key="excel_save_btn"):
        items = st.session_state.get("overall_items", [])
        if not items:
            st.error("明細がありません。品名行を1つ以上入力してください。")
            st.stop()

        # テンプレ存在チェック（独自フォーマット出力は不可）
        try:
            wb = load_workbook(str(TEMPLATE_BOOK), data_only=False)
            sn = set(wb.sheetnames)
            has_tpl = ("見積書0" in sn) and all(f"見積書{i}" in sn for i in range(1, 6))
        except Exception:
            has_tpl = False
        if not has_tpl:
            st.error("テンプレート『お見積書（明細）.xlsx』に『見積書0/見積書1〜5』が見つかりません。独自フォーマットでの出力は禁止されています。")
            st.stop()

        header = {
            "見積番号": st.session_state.estimate_no,
            "作成日":  date.today().strftime("%Y/%m/%d"),
            "得意先名": st.session_state.customer_name,
            "支店名":   st.session_state.branch_name,
            "営業所名": st.session_state.office_name,
            "担当者名": st.session_state.rep_name,
            "物件名":   st.session_state.project_name,
            "shipping": {
                "label": "運賃・梱包",
                "option": st.session_state.ship_option,
                "qty": 1, "unit": "式",
                "price": int(st.session_state.ship_price),
            },
        }

        try:
            # ロジック検証
            validate(items, header, rows_per_page=33, max_pages=5)

            out = osp.join(os.getcwd(), f"{st.session_state.get('file_title','見積')}_お見積書（明細）.xlsx")
            export_quotation_book_preserve(
                out, header, items,
                template_path=str(TEMPLATE_BOOK),
                header_sheet="見積書0",
                detail_sheets=[f"見積書{i}" for i in range(1, 6)],
                start_row=12, end_row=44,
            )

            st.success(f"Excelを保存しました：{out}")
            try:
                with open(out, "rb") as f:
                    st.download_button(
                        "ダウンロード", f.read(),
                        file_name=os.path.basename(out),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="excel_dl_btn",
                    )
            except Exception:
                pass
        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error("Excel出力でエラーが発生しました。")
            st.exception(e)


def main():
    _ensure_session()
    _customer_block()         # 得意先情報（J1/J3/A6/A7/A8/B17）
    _product_master_block()   # 製品情報（master.xlsx 連動）
    _product_free_block()     # 未登録時の手入力フォールバック
    _openings_block()         # 明細編集（入力順保持）
    _shipping_block()         # 運賃・梱包（見積書0の末尾）
    _save_block()             # 保存（テンプレ保持のみ）

if __name__ == "__main__":
    main()
