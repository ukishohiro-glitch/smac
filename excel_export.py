# -*- coding: utf-8 -*-
# excel_export.py — Excel転記（テンプレ保持のみ）
# - 見積書0: J1/J3/A6/A7/A8/B17 と 21〜44 に間口合計（空行なし）、直後に「運賃・梱包」
# - 見積書1〜5: 入力順で 12〜44（A=品名、F=数量、G=単位、H=単価）
# - 1間口<=33はページ跨ぎ不可、>33は跨ぎ可（呼び出し側で検証済みだが念のため防御）
from typing import List, Dict, Any, Sequence
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

def _to_int(x) -> int:
    try:
        if x is None: return 0
        if isinstance(x, int): return max(0, x)
        s = str(x).replace(",", "").strip()
        if not s: return 0
        return max(0, int(float(s)))
    except Exception:
        return 0

def _opening_total(items: List[Dict[str, Any]]) -> int:
    return sum(_to_int(it.get("数量", 0)) * _to_int(it.get("単価", 0)) for it in (items or []))

def _first_curtain_name(items: List[Dict[str, Any]]) -> str:
    if not items: return ""
    patterns = [
        "S・MAC",
        "エア・セーブ",  # 各型含む
        "ME(カーテン)", "ME（カーテン）", "ME(ｶｰﾃﾝ)", "ME（ｶｰﾃﾝ）",
        "カーテン",
    ]
    for it in items:
        nm = str(it.get("品名") or "")
        for p in patterns:
            if p in nm:
                return nm
    return str(items[0].get("品名") or "")

def _normalize_openings(overall: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return [{"name": str(op.get("name") or "").strip(),
             "items": list(op.get("items") or []),
             "teikeiku": list(op.get("teikeiku") or [])} for op in (overall or [])]

# ---- 見積書0
def _write_header(ws0: Worksheet, header: Dict[str, Any]):
    ws0["J1"].value = header.get("見積番号","")
    ws0["J3"].value = header.get("作成日","")
    ws0["A6"].value = header.get("得意先名","")
    ws0["A7"].value = f"{header.get('支店名','')} {header.get('営業所名','')}".strip()
    ws0["A8"].value = header.get("担当者名","")
    ws0["B17"].value = header.get("物件名","")

def _write_sheet0_totals(ws0: Worksheet, openings: List[Dict[str, Any]], header: Dict[str, Any]):
    row = 21
    for op in openings:
        items = op.get("items") or []
        ws0.cell(row=row, column=1).value = _first_curtain_name(items)   # A列
        ws0.cell(row=row, column=6).value = 1                            # F列
        ws0.cell(row=row, column=7).value = "式"                         # G列
        ws0.cell(row=row, column=8).value = _opening_total(items)        # H列
        row += 1
        if row > 44:
            raise ValueError("見積書0が 21〜44 行を超過しました。")

    ship = header.get("shipping") or {}
    price = _to_int(ship.get("price", 0))
    if price > 0:
        ws0.cell(row=row, column=1).value = f"{ship.get('label','運賃・梱包')}{ship.get('option','')}"
        ws0.cell(row=row, column=6).value = 1
        ws0.cell(row=row, column=7).value = "式"
        ws0.cell(row=row, column=8).value = price

# ---- 見積書1〜5
def _write_detail_row(ws: Worksheet, row: int, it: Dict[str, Any]):
    ws.cell(row=row, column=1).value = str(it.get("品名") or "")
    ws.cell(row=row, column=6).value = _to_int(it.get("数量", 0))
    ws.cell(row=row, column=7).value = str(it.get("単位") or "式")
    ws.cell(row=row, column=8).value = _to_int(it.get("単価", 0))

def _write_details(wb, openings: List[Dict[str, Any]], detail_sheets: Sequence[str],
                   start_row: int = 12, end_row: int = 44):
    rows_per_page = end_row - start_row + 1  # 33
    page_idx = 0
    ws = wb[detail_sheets[page_idx]]
    r = start_row

    def next_page():
        nonlocal page_idx, ws, r
        page_idx += 1
        if page_idx >= len(detail_sheets):
            raise ValueError("5ページを超えました。生成不可")
        ws = wb[detail_sheets[page_idx]]
        r = start_row

    for op in openings:
        items = list(op.get("items") or [])
        teikei = list(op.get("teikeiku") or [])
        need = len(items) + len(teikei)

        if need <= rows_per_page:
            if r + need - 1 > end_row:
                next_page()
            for it in items:
                _write_detail_row(ws, r, it); r += 1
            for s in teikei:
                ws.cell(row=r, column=1).value = str(s); r += 1
            if r > end_row:
                if r != end_row + 1:
                    while r <= end_row:
                        r += 1
                if r > end_row:
                    next_page()
        else:
            # 跨ぎ可
            rest_items = list(items)
            rest_teikei = list(teikei)

            def write_chunk(cap: int):
                nonlocal r
                written = 0
                while rest_items and written < cap:
                    _write_detail_row(ws, r, rest_items.pop(0))
                    r += 1; written += 1
                while rest_teikei and written < cap:
                    ws.cell(row=r, column=1).value = str(rest_teikei.pop(0))
                    r += 1; written += 1

            while rest_items or rest_teikei:
                if r > end_row:
                    next_page()
                cap = end_row - r + 1
                if cap <= 0:
                    next_page(); cap = end_row - r + 1
                write_chunk(cap)

# ---- 公開API（テンプレ保持のみ）
def export_quotation_book_preserve(out_path: str,
                                   header: Dict[str, Any],
                                   overall_openings: List[Dict[str, Any]],
                                   *,
                                   template_path: str,
                                   header_sheet: str = "見積書0",
                                   detail_sheets: Sequence[str] = None,
                                   start_row: int = 12,
                                   end_row: int = 44):
    if detail_sheets is None:
        detail_sheets = [f"見積書{i}" for i in range(1, 6)]

    wb = load_workbook(template_path)
    ws0 = wb[header_sheet]
    openings = _normalize_openings(overall_openings)

    _write_header(ws0, header)
    _write_sheet0_totals(ws0, openings, header)
    _write_details(wb, openings, detail_sheets, start_row=start_row, end_row=end_row)

    wb.save(out_path)
