# main.py - 完全版（全文書き換え / UI含むフル実装）

import os
import os.path as osp
import secrets
import re
from pathlib import Path
from datetime import datetime, date

import streamlit as st
from openpyxl import load_workbook

# ===== パス定義（__file__ 非対応環境でも安全） =====
try:
    APP_DIR = Path(__file__).parent
except NameError:
    APP_DIR = Path.cwd()
TEMPLATE_BOOK = APP_DIR / "お見積書（明細）.xlsx"
MASTER_BOOK   = APP_DIR / "master.xlsx"

# ===== Excel エクスポート（直 import に統一） =====
from excel_export import export_quotation_book_preserve, export_detail_xlsx_preserve


# ===== 見積番号：3709-xxxxx を一度だけ採番 =====
def _next_estimate_no() -> str:
    prefix = "3709-"
    used = st.session_state.setdefault("used_estnos", set())
    for _ in range(1000):
        rnd = f"{secrets.randbelow(100000):05d}"
        eno = prefix + rnd
        if eno not in used:
            used.add(eno)
            return eno
    return prefix + datetime.now().strftime("%H%M%S")[-5:]


# ===== S・MAC または エア・セーブ MA 型を含む → 梱包・運賃必須 =====
def _shipping_required(items: list[dict]) -> bool:
    for it in items or []:
        name = str(it.get("品名") or it.get("name") or it.get("product") or "")
        if "S・MAC" in name:
            return True
        if "エア・セーブ" in name and re.search(r"\bMA\b|MA型", name):
            return True
    return False


# ===== セッション初期化 =====
def _ensure_session():
    ss = st.session_state
    ss.setdefault("file_title", "見積")
    ss.setdefault("rep_name", "")
    if "estimate_no" not in ss:
        ss.estimate_no = _next_estimate_no()

    # 得意先情報
    ss.setdefault("customer_name", "")
    ss.setdefault("branch_name", "")
    ss.setdefault("office_name", "")
    ss.setdefault("project_name", "")

    # 明細（間口）構造
    # openings = [ { "name":"開口1", "items":[{品名,数量,単位,単価},...], "teikeiku":[...] }, ... ]
    if "openings" not in ss:
        ss.openings = [{
            "name": "開口1",
            "items": [{"品名":"", "数量":1, "単位":"式", "単価":0}],
            "teikeiku": []
        }]
    ss.setdefault("ship_option", "（路線便時間指定不可）")
    ss.setdefault("ship_price", 0)

    # overall_items は保存前に生成（UI編集中も更新しておく）
    ss.setdefault("overall_items", [])


# ===== ヘッダー（担当者名・見積番号） =====
def _header_section():
    st.title("お見積書作成システム")

    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.session_state.rep_name = st.text_input(
            "担当者名（半角英数2〜4）",
            value=st.session_state.rep_name,
            key="rep_name_input",
        )
        if st.session_state.rep_name and not re.fullmatch(r"[A-Za-z0-9]{2,4}", st.session_state.rep_name):
            st.error("担当者名は半角英数字2〜4文字で入力してください。")
    with col2:
        st.text_input("見積番号", st.session_state.estimate_no, key="estimate_no", disabled=True)
    with col3:
        st.text_input("保存ファイル名（接頭辞）", st.session_state.file_title, key="file_title")

    st.divider()


# ===== 得意先情報 =====
def _customer_section():
    st.markdown("### 得意先情報")
    c1, c2, c3 = st.columns([2, 2, 2])
    with c1:
        st.session_state.customer_name = st.text_input("得意先名", value=st.session_state.customer_name, key="customer_name")
    with c2:
        st.session_state.branch_name = st.text_input("支店名", value=st.session_state.branch_name, key="branch_name")
    with c3:
        st.session_state.office_name = st.text_input("営業所名", value=st.session_state.office_name, key="office_name")

    st.session_state.project_name = st.text_input("物件名（見積書0のB17）", value=st.session_state.project_name, key="project_name")
    st.divider()


# ===== 明細編集（間口単位・入力順） =====
def _openings_section():
    st.markdown("### 明細入力（間口ごと・入力順で転記）")
    # 間口一覧
    for oi, op in enumerate(st.session_state.openings):
        with st.expander(f"間口 {oi+1}：{op.get('name') or '(名称未設定)'}", expanded=True):
            # 間口名
            op["name"] = st.text_input("間口名", value=op.get("name",""), key=f"opening_name_{oi}")

            # 明細行エディタ
            st.write("明細行（A=品名 / F=数量 / G=単位 / H=単価）")
            # 各行
            remove_idx = None
            for ri, row in enumerate(op["items"]):
                cols = st.columns([5, 1, 1, 2, 0.8])
                with cols[0]:
                    row["品名"] = st.text_input("品名", value=str(row.get("品名","")), key=f"op{oi}_row{ri}_name")
                with cols[1]:
                    row["数量"] = st.number_input("数量", value=int(row.get("数量",1)), step=1, min_value=0, key=f"op{oi}_row{ri}_qty")
                with cols[2]:
                    row["単位"] = st.text_input("単位", value=str(row.get("単位","式")), key=f"op{oi}_row{ri}_unit")
                with cols[3]:
                    row["単価"] = st.number_input("単価（円）", value=int(row.get("単価",0)), step=100, min_value=0, key=f"op{oi}_row{ri}_price")
                with cols[4]:
                    if st.button("−", key=f"op{oi}_row{ri}_del"):
                        remove_idx = ri
                st.write("")  # 下部マージン
            if remove_idx is not None and 0 <= remove_idx < len(op["items"]):
                del op["items"][remove_idx]

            # 行追加
            if st.button("＋ 行を追加", key=f"op{oi}_addrow"):
                op["items"].append({"品名":"", "数量":1, "単位":"式", "単価":0})

            st.write("---")
            # 定型句：末尾の A 列に出力
            teikei_text = "\n".join(op.get("teikeiku") or [])
            teikei_text = st.text_area("定型句（1行=1文、間口末尾の A 列に出力）", value=teikei_text, key=f"op{oi}_teikei", height=80)
            op["teikeiku"] = [s for s in teikei_text.splitlines() if s.strip()]

            # 間口削除
            cols_del = st.columns([6,1])
            with cols_del[1]:
                if st.button("× この間口を削除", key=f"op{oi}_del"):
                    st.session_state.openings.pop(oi)
                    st.experimental_rerun()

    # 間口追加
    if st.button("＋ 間口を追加", key="opening_add"):
        st.session_state.openings.append({
            "name": f"開口{len(st.session_state.openings)+1}",
            "items": [{"品名":"", "数量":1, "単位":"式", "単価":0}],
            "teikeiku": []
        })

    st.divider()

    # overall_items を更新（形式1に統一）
    st.session_state.overall_items = [{
        "name": op.get("name",""),
        "items": list(op.get("items") or []),
        "teikeiku": list(op.get("teikeiku") or []),
    } for op in st.session_state.openings]


# ===== 運賃・梱包（見積書0の末尾に記載・明細には載せない） =====
def _shipping_section():
    st.markdown("### 運賃・梱包（見積書0の末尾に記載・明細には載せません）")
    ship_c1, ship_c2 = st.columns([2, 2])
    with ship_c1:
        st.session_state.ship_option = st.radio(
            "配送条件",
            options=["（路線便時間指定不可）", "（現場搬入時間指定可）"],
            index=0 if st.session_state.ship_option == "（路線便時間指定不可）" else 1,
            key="ship_option",
            horizontal=False,
        )
    with ship_c2:
        st.session_state.ship_price = st.number_input(
            "運賃・梱包 金額（円）",
            min_value=0,
            step=100,
            value=int(st.session_state.ship_price),
            key="ship_price",
        )
    st.divider()


# ===== 保存（唯一の保存ボタン） =====
def _save_section():
    st.markdown("### 保存")
    s1, s2 = st.columns([1, 2])
    with s2:
        if st.button("Excel保存（お見積書（明細））", key="excel_save_btn"):
            overall_items = st.session_state.get("overall_items", [])
            if not overall_items:
                st.error("明細がありません。")
                st.stop()

            # 梱包・運賃 必須チェック
            flat_items = []
            for op in overall_items:
                flat_items.extend(op.get("items", []))
            if _shipping_required(flat_items) and int(st.session_state.get("ship_price", 0)) <= 0:
                st.error("S・MAC または エア・セーブ MA 型を含むため、運賃・梱包の金額入力が必須です。")
                st.stop()

            # ヘッダ辞書（見積書0へ）
            header = {
                "見積番号": st.session_state.estimate_no,
                "作成日": date.today().strftime("%Y/%m/%d"),
                "得意先名": st.session_state.customer_name,
                "支店名": st.session_state.branch_name,
                "営業所名": st.session_state.office_name,
                "担当者名": st.session_state.rep_name,
                "物件名": st.session_state.project_name,
                # 見積書0の末尾に『空行なしで連続』追記（明細には載せない）
                "shipping": {
                    "label": "運賃・梱包",
                    "option": st.session_state.ship_option,
                    "qty": 1,
                    "unit": "式",
                    "price": int(st.session_state.ship_price),
                },
            }

            file_title = st.session_state.get("file_title", "見積")
            out = osp.join(os.getcwd(), f"{file_title}_お見積書（明細）.xlsx")
            tpl = str(TEMPLATE_BOOK) if TEMPLATE_BOOK.exists() else str(APP_DIR / "お見積書（明細）.xlsx")

            try:
                # 新テンプレ（見積書0/1〜5）判定
                use_new = False
                if tpl and osp.exists(tpl):
                    try:
                        _wb_check = load_workbook(tpl, data_only=False)
                        sn = set(_wb_check.sheetnames)
                        use_new = ("見積書0" in sn) and all(f"見積書{i}" in sn for i in range(1, 6))
                    except Exception:
                        use_new = False

                if use_new and export_quotation_book_preserve is not None:
                    # 見積書0：ヘッダー＆間口合計（空行なし連続）→直後に運賃・梱包
                    # 見積書1〜5：明細（12〜44行、11行目は触らない／A=品名 F=数量 G=単位 H=単価）
                    export_quotation_book_preserve(
                        out, header, overall_items,
                        template_path=tpl,
                        header_sheet="見積書0",
                        detail_sheets=[f"見積書{i}" for i in range(1, 6)],
                        start_row=12, end_row=44,
                    )
                else:
                    # 旧テンプレ互換（単一シート／33行/頁／5頁制限）
                    if export_detail_xlsx_preserve is None:
                        raise ImportError("旧テンプレ互換の export_detail_xlsx_preserve が見つかりません。")
                    export_detail_xlsx_preserve(
                        out, header, overall_items,
                        template_path=tpl if (tpl and osp.exists(tpl)) else None,
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
                # 例：5ページ超過「5ページを超えました。生成不可」
                st.error(str(e))
            except Exception as e:
                st.error("Excel出力でエラーが発生しました。")
                st.exception(e)


def main():
    st.set_page_config(layout="wide")
    _ensure_session()
    _header_section()
    _customer_section()
    _openings_section()
    _shipping_section()
    _save_section()


if __name__ == "__main__":
    main()
# excel_export.py - テンプレ罫線/結合を維持した転記（見積書0/1〜5対応）

from __future__ import annotations
from typing import List, Dict, Any
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


# ========== 内部ユーティリティ ==========

def _to_int(x) -> int:
    try:
        return int(x)
    except Exception:
        try:
            return int(float(x))
        except Exception:
            return 0


def _first_curtain_name(items: List[Dict[str, Any]]) -> str:
    """間口内で最初に登場したカーテン種別（S・MAC／エア・セーブ各型／ME(カーテン)）を採用"""
    if not items:
        return ""
    keywords = ["S・MAC", "エア・セーブ", "ME", "カーテン"]
    for it in items:
        name = str(it.get("品名") or it.get("name") or it.get("product") or "")
        for kw in keywords:
            if kw in name:
                return name
    return str(items[0].get("品名") or items[0].get("name") or items[0].get("product") or "")


def _opening_total(items: List[Dict[str, Any]]) -> int:
    """間口合計：各明細 (数量×単価) の合計"""
    total = 0
    for it in items:
        qty = _to_int(it.get("数量") or it.get("qty") or 0)
        unit_price = _to_int(it.get("単価") or it.get("unit_price") or 0)
        total += qty * unit_price
    return total


def _normalize_openings(overall_items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    overall_items を「間口」配列に正規化。
      1) [{ "name": "...", "items": [...], "teikeiku": [...] }, ...]
      2) フラット配列に "開口"/"間口" → 連続同名でグループ化
      3) どれでもない → 全体を1間口
    """
    if not overall_items:
        return []

    if isinstance(overall_items[0], dict) and "items" in overall_items[0]:
        openings = []
        for op in overall_items:
            openings.append({
                "name": op.get("name") or op.get("開口") or op.get("間口") or "",
                "items": list(op.get("items") or []),
                "teikeiku": list(op.get("teikeiku") or []),
            })
        return openings

    has_key = any(("開口" in it) or ("間口" in it) for it in overall_items if isinstance(it, dict))
    if has_key:
        openings = []
        current_name = None
        bucket: List[Dict[str, Any]] = []
        for it in overall_items:
            name = it.get("開口") or it.get("間口") or ""
            if current_name is None:
                current_name = name
            if name != current_name:
                openings.append({"name": current_name or "", "items": bucket, "teikeiku": []})
                bucket = []
                current_name = name
            bucket.append(it)
        if bucket:
            openings.append({"name": current_name or "", "items": bucket, "teikeiku": []})
        return openings

    return [{"name": "", "items": overall_items, "teikeiku": []}]


# ========== 見積書0：ヘッダー & 集約 ==========

def _write_header(ws0: Worksheet, header: Dict[str, Any]) -> None:
    # 指定：見積番号:J1 / 作成日:J3 / 得意先名:A6 / 支店名 営業所名:A7（間は半角スペース）/ 担当者名:A8 / 物件名:B17
    ws0["J1"].value = header.get("見積番号", "")
    ws0["J3"].value = header.get("作成日", "")
    ws0["A6"].value = header.get("得意先名", "")
    b = header.get("支店名", "")
    o = header.get("営業所名", "")
    ws0["A7"].value = (str(b) + " " + str(o)).strip()
    ws0["A8"].value = header.get("担当者名", "")
    ws0["B17"].value = header.get("物件名", "")


def _write_sheet0_totals(ws0: Worksheet, openings: List[Dict[str, Any]], header: Dict[str, Any]) -> None:
    """
    A=品名（カーテン品名）/ F=数量(1) / G=単位(式) / H=単価(=間口合計)
    21〜44行に『空行なしで連続』で記載。直後に「運賃・梱包」を追記（空行なし）。
    """
    row = 21  # 指定：21〜44
    for op in openings:
        items = op.get("items", [])
        curtain = _first_curtain_name(items)
        total = _opening_total(items)

        ws0.cell(row=row, column=1).value = curtain      # A
        ws0.cell(row=row, column=6).value = 1            # F
        ws0.cell(row=row, column=7).value = "式"         # G
        ws0.cell(row=row, column=8).value = total        # H

        row += 1
        if row > 44:
            raise ValueError("見積書0の行数が 21〜44 を超えました。")

    # shipping を直後に追記（空行なし）
    ship = header.get("shipping")
    if ship and _to_int(ship.get("price", 0)) > 0:
        label = str(ship.get("label", "運賃・梱包")) + str(ship.get("option", ""))
        ws0.cell(row=row, column=1).value = label
        ws0.cell(row=row, column=6).value = 1
        ws0.cell(row=row, column=7).value = "式"
        ws0.cell(row=row, column=8).value = _to_int(ship.get("price", 0))


# ========== 見積書1〜5：明細（入力順） ==========

def _write_detail_row(ws: Worksheet, row: int, it: Dict[str, Any]) -> None:
    """A=品名 / F=数量 / G=単位 / H=単価（11行目は触らない・12行目から）"""
    name = str(it.get("品名") or it.get("name") or it.get("product") or "")
    qty = it.get("数量") if it.get("数量") is not None else it.get("qty")
    unit = it.get("単位") if it.get("単位") is not None else it.get("unit")
    price = it.get("単価") if it.get("単価") is not None else it.get("unit_price")

    ws.cell(row=row, column=1).value = name
    ws.cell(row=row, column=6).value = qty
    ws.cell(row=row, column=7).value = unit
    ws.cell(row=row, column=8).value = price


def _write_details(
    wb,
    openings: List[Dict[str, Any]],
    detail_sheets: List[str],
    start_row: int = 12,
    end_row: int = 44,
) -> None:
    """
    - 表示は入力順
    - 11行目は触らない → start_row=12 固定
    - 1ページ=33行（12〜44）
    - 1間口が33行以内なら『同一ページ内に収める』（不足ならページ送り）
    - 1間口が33行を超える場合のみページ跨ぎOK
    - 最大5ページ、超えたら ValueError("5ページを超えました。生成不可")
    - 定型句は各間口末尾に A 列
    """
    rows_per_page = end_row - start_row + 1
    page_max = len(detail_sheets)
    page_idx = 0
    ws = wb[detail_sheets[page_idx]]
    r = start_row

    def ensure_next_page():
        nonlocal page_idx, ws, r
        page_idx += 1
        if page_idx >= page_max:
            raise ValueError("5ページを超えました。生成不可")
        ws = wb[detail_sheets[page_idx]]
        r = start_row

    for op in openings:
        items = list(op.get("items", []))
        teikeiku = list(op.get("teikeiku", []))
        need = len(items) + len(teikeiku)

        if need <= rows_per_page:
            # 今のページに入らなければページ送りして丸ごと収める
            if r + need - 1 > end_row:
                ensure_next_page()
            for it in items:
                _write_detail_row(ws, r, it)
                r += 1
            for s in teikeiku:
                ws.cell(row=r, column=1).value = str(s)
                r += 1
        else:
            # 33行超：ページ跨ぎで連続記載
            idx = 0
            while idx < len(items):
                if r > end_row:
                    ensure_next_page()
                _write_detail_row(ws, r, items[idx])
                r += 1
                idx += 1
            for s in teikeiku:
                if r > end_row:
                    ensure_next_page()
                ws.cell(row=r, column=1).value = str(s)
                r += 1


# ========== 公開関数：新テンプレ（見積書0 + 見積書1〜5） ==========

def export_quotation_book_preserve(
    out_path: str,
    header: Dict[str, Any],
    overall_items: List[Dict[str, Any]],
    *,
    template_path: str,
    header_sheet: str = "見積書0",
    detail_sheets: List[str] | None = None,
    start_row: int = 12,
    end_row: int = 44,
) -> None:
    if detail_sheets is None:
        detail_sheets = [f"見積書{i}" for i in range(1, 6)]

    wb = load_workbook(template_path)
    if header_sheet not in wb.sheetnames:
        raise ValueError(f"テンプレートに {header_sheet} がありません。")
    for s in detail_sheets:
        if s not in wb.sheetnames:
            raise ValueError(f"テンプレートに {s} がありません。")

    # 見積書0：ヘッダー & 集約
    ws0 = wb[header_sheet]
    _write_header(ws0, header)
    openings = _normalize_openings(overall_items)
    _write_sheet0_totals(ws0, openings, header)

    # 見積書1〜5：明細
    _write_details(wb, openings, detail_sheets, start_row=start_row, end_row=end_row)

    wb.save(out_path)


# ========== 公開関数：旧テンプレ互換（単一シート） ==========

def export_detail_xlsx_preserve(
    out_path: str,
    header: Dict[str, Any],
    overall_items: List[Dict[str, Any]],
    *,
    template_path: str | None,
    ws_name: str = "お見積書（明細）",
    start_row: int = 12,
    max_rows: int = 33,
) -> None:
    """
    旧テンプレ互換：単一シートに 33 行/ページで積み上げ、5ページ超で停止。
    shipping は明細に載せない（ユーザー指示）。
    """
    if not template_path:
        raise ValueError("旧テンプレ互換には template_path が必要です。")
    wb = load_workbook(template_path)
    if ws_name not in wb.sheetnames:
        raise ValueError(f"テンプレートに {ws_name} がありません。")

    ws = wb[ws_name]
    ws["J1"].value = header.get("見積番号", "")
    ws["J3"].value = header.get("作成日", "")

    page = 0
    row = start_row
    rows_per_page = max_rows
    pages_limit = 5

    openings = _normalize_openings(overall_items)

    def ensure_next_page():
        nonlocal page, row
        page += 1
        if page >= pages_limit:
            raise ValueError("5ページを超えました。生成不可")
        row = start_row + page * rows_per_page

    for op in openings:
        items = list(op.get("items", []))
        teikeiku = list(op.get("teikeiku", []))
        need = len(items) + len(teikeiku)

        while (row - (start_row + page * rows_per_page)) + need > rows_per_page:
            ensure_next_page()

        for it in items:
            _write_detail_row(ws, row, it)
            row += 1

        for s in teikeiku:
            ws.cell(row=row, column=1).value = str(s)
            row += 1

    wb.save(out_path)
