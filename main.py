# -*- coding: utf-8 -*-
# main.py - UI層（Streamlit）
# できること：
# - master.xlsxを読み込み、カテゴリ→検索→部材選択→単価/単位を自動補完（手入力も可）
# - 入力順で明細を保持（間口ごと）
# - 11行目は触らない（明細は12〜44）/ 見積書0は21〜44
# - 1間口33行以内はページ跨ぎ禁止 / 33行超は跨ぎOK / 5ページ超はエラー
# - S・MAC or エア・セーブMA含有時は「運賃・梱包」金額必須
# - 見積番号は 3709-xxxxx 固定（再生成なし）
# - Excel転記はテンプレの罫線・結合を維持（見積書0/1〜5へ転記）

import os
import os.path as osp
import re
import secrets
from pathlib import Path
from datetime import date, datetime
from typing import Dict, Any, List

import streamlit as st
from openpyxl import load_workbook

# パス解決（__file__が無い環境も考慮）
try:
    APP_DIR = Path(__file__).parent
except NameError:
    APP_DIR = Path.cwd()
TEMPLATE_BOOK = APP_DIR / "お見積書（明細）.xlsx"
MASTER_BOOK   = APP_DIR / "master.xlsx"

# 業務ロジック / Excel出力
from quote_logic import clean_openings, plan_paging, validate
from excel_export import export_quotation_book_preserve, export_detail_xlsx_preserve


# =========================
#   共通ユーティリティ
# =========================
def _next_estimate_no() -> str:
    """見積番号は 3709-xxxxx（5桁）。重複はセッション内で回避。"""
    used = st.session_state.setdefault("used_estnos", set())
    for _ in range(2000):
        rnd = f"{secrets.randbelow(100000):05d}"
        eno = "3709-" + rnd
        if eno not in used:
            used.add(eno)
            return eno
    return "3709-" + datetime.now().strftime("%H%M%S")[-5:]


# =========================
#   master.xlsx ロード
# =========================
def _normalize_header_map(headers: List[str]) -> Dict[str, int]:
    idx = {str(h or "").strip(): i for i, h in enumerate(headers)}
    def find(*cands):
        for c in cands:
            for h, i in idx.items():
                if h == c or (c and c in h):
                    return i
        return None
    return {
        "品名": find("品名", "商品名", "名称", "製品名", "品番"),
        "単位": find("単位", "Unit"),
        "単価": find("基準単価", "単価", "上代", "価格", "金額"),
        "型番": find("型番", "型式", "コード"),
    }

def load_master_book(path: Path) -> Dict[str, List[Dict[str, Any]]]:
    """master.xlsx を {シート名: [{品名, 単位, 単価, 型番}, ...]} にして返す。1行目ヘッダ。"""
    catalog: Dict[str, List[Dict[str, Any]]] = {}
    if not path.exists():
        return catalog
    wb = load_workbook(str(path), data_only=True, read_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        headers = [str(x or "").strip() for x in rows[0]]
        colmap = _normalize_header_map(headers)
        if colmap["品名"] is None:
            continue
        items: List[Dict[str, Any]] = []
        for r in rows[1:]:
            if not r:
                continue
            name = str(r[colmap["品名"]] if colmap["品名"] is not None and colmap["品名"] < len(r) else "").strip()
            if not name:
                continue
            unit = str(r[colmap["単位"]] if colmap["単位"] is not None and colmap["単位"] < len(r) else "式").strip() or "式"
            raw_price = r[colmap["単価"]] if colmap["単価"] is not None and colmap["単価"] < len(r) else 0
            try:
                price = int(float(str(raw_price).replace(",", "").strip()))
                if price < 0:
                    price = 0
            except Exception:
                price = 0
            model = str(r[colmap["型番"]] if colmap["型番"] is not None and colmap["型番"] < len(r) else "").strip()
            items.append({"品名": name, "単位": unit, "単価": price, "型番": model})
        if items:
            catalog[ws.title] = items
    return catalog


# =========================
#   セッション初期化
# =========================
def _ensure_session():
    ss = st.session_state
    ss.setdefault("file_title", "見積")
    if "estimate_no" not in ss:
        ss.estimate_no = _next_estimate_no()

    # ヘッダー入力
    ss.setdefault("customer_name", "")
    ss.setdefault("branch_name", "")
    ss.setdefault("office_name", "")
    ss.setdefault("rep_name", "")
    ss.setdefault("project_name", "")

    # 間口/定型句
    if "openings" not in ss:
        ss.openings = [{
            "name": "開口1",
            "items": [{"品名": "", "数量": 1, "単位": "式", "単価": 0}],
            "teikeiku": []
        }]
    ss.setdefault("teikei_map", {})  # {opening_idx: [str,...]}

    # 運賃・梱包
    ss.setdefault("ship_option", "（路線便時間指定不可）")
    ss.setdefault("ship_price", 0)

    # マスタ
    ss.setdefault("catalog", load_master_book(MASTER_BOOK))
    ss.setdefault("search_query", "")


# =========================
#   UI セクション
# =========================
def _header_section():
    st.set_page_config(layout="wide", page_title="お見積書作成")
    st.title("お見積書作成システム")

    c1, c2, c3 = st.columns([1.2, 0.8, 1.0])
    with c1:
        st.text_input("見積番号（固定：3709-xxxxx）", value=st.session_state.estimate_no, key="estimate_no", disabled=True)
    with c2:
        st.session_state.file_title = st.text_input("保存ファイル名（接頭辞）", value=st.session_state.file_title, key="file_title")
    with c3:
        st.text_input("作成日", value=date.today().strftime("%Y/%m/%d"), key="created_disp", disabled=True)

    d1, d2, d3, d4 = st.columns([2, 2, 2, 2])
    with d1:
        st.session_state.customer_name = st.text_input("得意先名（見積書0:A6）", value=st.session_state.customer_name, key="customer_name")
    with d2:
        st.session_state.branch_name = st.text_input("支店名（見積書0:A7 左）", value=st.session_state.branch_name, key="branch_name")
    with d3:
        st.session_state.office_name = st.text_input("営業所名（見積書0:A7 右）", value=st.session_state.office_name, key="office_name")
    with d4:
        rep = st.text_input("担当者名（半角英数2〜4）", value=st.session_state.rep_name, key="rep_name")
        if rep and not re.fullmatch(r"[A-Za-z0-9]{2,4}", rep):
            st.error("担当者名は半角英数字2〜4文字で入力してください。")
        st.session_state.rep_name = rep

    st.text_input("物件名（見積書0:B17）", value=st.session_state.project_name, key="project_name")
    st.divider()


def _picker_section():
    st.markdown("### 部材マスタから追加（master.xlsx）")
    catalog = st.session_state.catalog
    if not catalog:
        st.info("master.xlsx が見つからないか読み込めません。手入力で続行できます。")
        return

    cat_names = sorted(catalog.keys())
    cols = st.columns([2, 2, 3, 1, 1])
    with cols[0]:
        cat = st.selectbox("カテゴリ（シート名）", options=cat_names, key="cat_sel")
    items = catalog.get(cat, [])
    with cols[1]:
        st.session_state.search_query = st.text_input("キーワード検索", value=st.session_state.search_query, key="search_query_in")
    q = st.session_state.search_query.strip()
    filtered = [it for it in items if q.lower() in (it["品名"] + " " + (it.get("型番") or "")).lower()] if q else items

    name_list = [f'{it["品名"]} 〔{it.get("型番","") or "-"} / {it["単価"]}円 / {it["単位"]}〕' for it in filtered] or ["（該当なし）"]
    with cols[2]:
        sel = st.selectbox("部材を選択", options=name_list, key="part_sel")
    with cols[3]:
        qty = st.number_input("数量", min_value=1, step=1, value=1, key="pick_qty")
    with cols[4]:
        tgt_opening_idx = st.number_input("挿入先 間口No", min_value=1, step=1, value=1, key="pick_opening_idx")

    def _add_selected():
        if not filtered:
            return
        idx = name_list.index(sel) if sel in name_list else -1
        if idx < 0 or idx >= len(filtered):
            return
        item = dict(filtered[idx])  # 品名/単位/単価/型番
        item["数量"] = qty
        # 間口に追加（不足分は拡張）
        while len(st.session_state.openings) < tgt_opening_idx:
            st.session_state.openings.append({"name": f"開口{len(st.session_state.openings)+1}", "items": [], "teikeiku": []})
        st.session_state.openings[tgt_opening_idx - 1]["items"].append(
            {"品名": item["品名"], "数量": item["数量"], "単位": item["単位"], "単価": item["単価"]}
        )

    st.button("＋ この部材を追加", on_click=_add_selected, key="add_from_master_btn")
    st.divider()


def _openings_section():
    st.markdown("### 明細入力（間口ごと・入力順で転記） / 11行目は触らない → 12〜44行")
    for oi, op in enumerate(st.session_state.openings):
        with st.expander(f"間口 {oi+1}：{op.get('name') or '(名称未設定)'}", expanded=True):
            op["name"] = st.text_input("間口名", value=op.get("name", ""), key=f"opening_name_{oi}")
            st.caption("明細行（A=品名 / F=数量 / G=単位 / H=単価）")

            remove_idx = None
            for ri, row in enumerate(op["items"]):
                cols = st.columns([5, 1, 1, 2, 0.9])
                with cols[0]:
                    row["品名"] = st.text_input("品名", value=str(row.get("品名", "")), key=f"op{oi}_row{ri}_name")
                with cols[1]:
                    row["数量"] = st.number_input("数量", value=int(row.get("数量", 1)), step=1, min_value=0, key=f"op{oi}_row{ri}_qty")
                with cols[2]:
                    row["単位"] = st.text_input("単位", value=str(row.get("単位", "式")), key=f"op{oi}_row{ri}_unit")
                with cols[3]:
                    row["単価"] = st.number_input("単価（円）", value=int(row.get("単価", 0)), step=100, min_value=0, key=f"op{oi}_row{ri}_price")
                with cols[4]:
                    if st.button("−", key=f"op{oi}_row{ri}_del"):
                        remove_idx = ri
                st.write("")
            if remove_idx is not None and 0 <= remove_idx < len(op["items"]):
                del op["items"][remove_idx]
            if st.button("＋ 行を追加（手入力）", key=f"op{oi}_addrow"):
                op["items"].append({"品名": "", "数量": 1, "単位": "式", "単価": 0})

            st.write("---")
            teikei_text = "\n".join(op.get("teikeiku") or [])
            teikei_text = st.text_area("定型句（1行=1文、間口末尾の A 列に出力）", value=teikei_text, key=f"op{oi}_teikei", height=80)
            op["teikeiku"] = [s for s in teikei_text.splitlines() if s.strip()]

            cols_del = st.columns([6, 1])
            with cols_del[1]:
                if st.button("× この間口を削除", key=f"op{oi}_del") and len(st.session_state.openings) > 1:
                    st.session_state.openings.pop(oi)
                    st.experimental_rerun()

    if st.button("＋ 間口を追加", key="opening_add"):
        st.session_state.openings.append({
            "name": f"開口{len(st.session_state.openings)+1}",
            "items": [{"品名": "", "数量": 1, "単位": "式", "単価": 0}],
            "teikeiku": []
        })

    st.divider()

    # クリーニング（空行・数値型）
    st.session_state.overall_items = clean_openings([
        {"name": op.get("name", ""), "items": list(op.get("items") or []), "teikeiku": list(op.get("teikeiku") or [])}
        for op in st.session_state.openings
    ])

    # 概算ページ数
    if st.session_state.overall_items:
        try:
            est_pages = plan_paging(st.session_state.overall_items, rows_per_page=33)
            if est_pages > 5:
                st.warning(f"この入力だと明細ページは **約 {est_pages} ページ**（上限 5 ページ）。保存時にエラーになります。")
        except Exception:
            pass


def _shipping_section():
    st.markdown("### 運賃・梱包（見積書0の末尾に記載・明細には載せません）")
    c1, c2 = st.columns([2, 2])
    with c1:
        st.session_state.ship_option = st.radio(
            "配送条件",
            options=["（路線便時間指定不可）", "（現場搬入時間指定可）"],
            index=0 if st.session_state.ship_option == "（路線便時間指定不可）" else 1,
            key="ship_option",
            horizontal=True
        )
    with c2:
        st.session_state.ship_price = st.number_input(
            "運賃・梱包 金額（円）", min_value=0, step=100, value=int(st.session_state.ship_price), key="ship_price"
        )
    st.divider()


def _save_section():
    st.markdown("### 保存")
    if st.button("Excel保存（お見積書（明細））", key="excel_save_btn"):
        overall_items = st.session_state.get("overall_items", [])
        if not overall_items:
            st.error("明細がありません。間口の品名行を1つ以上入力してください。")
            st.stop()

        header = {
            "見積番号": st.session_state.estimate_no,                # 見積書0: J1
            "作成日":  date.today().strftime("%Y/%m/%d"),           # 見積書0: J3
            "得意先名": st.session_state.customer_name,             # 見積書0: A6
            "支店名":   st.session_state.branch_name,               # 見積書0: A7（左）
            "営業所名": st.session_state.office_name,               # 見積書0: A7（右、半角スペース）
            "担当者名": st.session_state.rep_name,                  # 見積書0: A8
            "物件名":   st.session_state.project_name,              # 見積書0: B17
            "shipping": {
                "label": "運賃・梱包",
                "option": st.session_state.ship_option,
                "qty": 1, "unit": "式",
                "price": int(st.session_state.ship_price),
            },
        }

        try:
            # 厳格バリデーション
            validate(overall_items, header, rows_per_page=33, max_pages=5)

            # 出力パス
            file_title = st.session_state.get("file_title", "見積")
            out = osp.join(os.getcwd(), f"{file_title}_お見積書（明細）.xlsx")
            tpl = str(TEMPLATE_BOOK)

            # テンプレ判定
            use_new = False
            try:
                wb = load_workbook(tpl, data_only=False)
                sn = set(wb.sheetnames)
                use_new = ("見積書0" in sn) and all(f"見積書{i}" in sn for i in range(1, 6))
            except Exception:
                use_new = False

            if use_new:
                export_quotation_book_preserve(
                    out, header, overall_items,
                    template_path=tpl,
                    header_sheet="見積書0",
                    detail_sheets=[f"見積書{i}" for i in range(1, 6)],
                    start_row=12, end_row=44,
                )
            else:
                # 旧テンプレ互換：1シートに 33行/頁 で追記（※テンプレ構成が違う場合の保険）
                export_detail_xlsx_preserve(
                    out, header, overall_items,
                    template_path=tpl,
                    ws_name="お見積書（明細）",
                    start_row=12, max_rows=33,
                )

            st.success(f"Excelを保存しました：{out}")
            with open(out, "rb") as f:
                st.download_button(
                    "ダウンロード", f.read(),
                    file_name=os.path.basename(out),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="excel_dl_btn",
                )
        except ValueError as e:
            st.error(str(e))
        except Exception as e:
            st.error("Excel出力でエラーが発生しました。")
            st.exception(e)


def main():
    _ensure_session()
    _header_section()
    _picker_section()
    _openings_section()
    _shipping_section()
    _save_section()

if __name__ == "__main__":
    main()
