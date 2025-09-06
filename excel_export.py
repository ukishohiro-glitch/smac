# -*- coding: utf-8 -*-
"""
excel_export.py — Excel転記ユーティリティ（テンプレ書式を完全維持）

要件：
- GPT独自の新規フォーマットを作らない。
- 既存テンプレートの罫線・セル結合・フォント・色・列幅・数式を壊さず、値だけを置き換える。
- 明細は「間口ごとに1行空ける」ロジックをサポート。
- 明細表示は **12行目〜44行目** を使用（＝最大33行）。
- ヘッダは **見積書0** シート（J1/J3/A6/A7/A8/B17）に転記、明細は **見積書1〜5** に転記。
- 明細列：A=品名, F=数量, G=単位, H=単価。
- 11行目は触らない。
- 定型句は間口末尾にA列。
- 間口合計は「見積書0」に転記（21〜44行）。カーテンの品名のみA列、数量F=1, 単位G=式, 単価H=合計。

本モジュールは main.py から import して使用します。
"""
from __future__ import annotations
import os.path as osp
from typing import List, Dict, Optional
from openpyxl import load_workbook

# =============================================
#  新様式：見積書0（ヘッダ/物件/間口合計）＋見積書1〜5（明細）に転記
# =============================================

def export_quotation_book_preserve(
    out_path: str,
    header: Dict,
    items: List[Dict],
    *,
    template_path: str,
    header_sheet: str = "見積書0",
    detail_sheets: Optional[List[str]] = None,
    start_row: int = 12,
    end_row: int = 44,
) -> None:
    if not (template_path and osp.exists(template_path)):
        raise FileNotFoundError("テンプレートファイルが見つかりません: " + str(template_path))

    wb = load_workbook(template_path, data_only=False)
    if header_sheet not in wb.sheetnames:
        raise ValueError(f"テンプレートにシート『{header_sheet}』がありません。")
    ws0 = wb[header_sheet]

    # --- ヘッダ転記 ---
    ws0["J1"].value = header.get("estimate_no", "")
    ws0["J3"].value = header.get("date", "")
    ws0["A6"].value = header.get("customer_name", "")
    b = header.get("branch_name", "") or ""
    o = header.get("office_name", "") or ""
    ws0["A7"].value = (b + " " + o).strip()
    ws0["A8"].value = header.get("person_name", "")
    ws0["B17"].value = header.get("project_name", "")

    # --- 明細シート群 ---
    if detail_sheets is None:
        detail_sheets = [f"見積書{i}" for i in range(1, 6)]
    ws_list = [wb[name] for name in detail_sheets if name in wb.sheetnames]
    if not ws_list:
        raise ValueError("テンプレートに『見積書1〜5』が見つかりません。")

    # 明細クリア（A/F/G/H のみ）
    for ws in ws_list:
        for r in range(start_row, end_row + 1):
            ws[f"A{r}"].value = None
            ws[f"F{r}"].value = None
            ws[f"G{r}"].value = None
            ws[f"H{r}"].value = None

    # === 明細を書き込み（定型句は間口末尾、間口間に1行空け） ===
    from collections import defaultdict
    group = defaultdict(list)
    for it in (items or []):
        group[it.get("_open")].append(it)

    def split_memo(seq):
        normal, memo = [], []
        for x in seq:
            kind = x.get("種別", "")
            name = x.get("品名", "")
            if kind == "メモ" or name == "（定型文）":
                memo.append(x)
            else:
                normal.append(x)
        return normal, memo

    sheet_idx = 0
    r = start_row
    rows_per_sheet = end_row - start_row + 1

    def ensure_next_sheet():
        nonlocal sheet_idx, r
        sheet_idx += 1
        if sheet_idx >= len(ws_list):
            raise ValueError("5ページを超えました。生成不可")
        r = start_row
        return True

    first_on_sheet = True
    # 開口単位でページをまたがない。ただし1開口が33行を超える場合はページをまたいで出力可。
    for op in sorted(group.keys(), key=lambda x: (x is None, x)):
        seq = group[op]
        normal, memo = split_memo(seq)
        ordered = normal + memo  # 末尾に定型句
        rows_needed = len(ordered) + (0 if first_on_sheet else 1)  # 間口間の空行を含む
        remaining = end_row - r + 1
        # 1開口が1ページに収まる場合で、残り行に入らないなら次ページに送る
        if len(ordered) <= rows_per_sheet and rows_needed > remaining:
            ensure_next_sheet()
            first_on_sheet = True
            remaining = rows_per_sheet
        # 間口間の空行（同一シートで先頭でない場合のみ）
        if not first_on_sheet:
            if r > end_row:
                ensure_next_sheet()
            else:
                r += 1
        ws = ws_list[sheet_idx]
        # 出力（1開口が33行を超える場合はページをまたいで継続）
        for it in ordered:
            if r > end_row:
                ensure_next_sheet()
                ws = ws_list[sheet_idx]
            name = it.get("品名", "")
            qty  = it.get("数量", "")
            unit = it.get("単位", "")
            price= it.get("単価", "")
            kind = it.get("種別", "")
            note = it.get("備考", "")
            # 列マッピング：A=品名, F=数量, G=単位, H=単価
            if kind == "メモ" or name == "（定型文）":
                ws[f"A{r}"].value = note  # 定型句はA列に備考のみ
            else:
                ws[f"A{r}"].value = name
                ws[f"F{r}"].value = qty
                ws[f"G{r}"].value = unit
                ws[f"H{r}"].value = price
            r += 1
        first_on_sheet = False

    # --- 間口合計を見積書0に転記（21〜44行） ---
    CURTAIN_KINDS = {"S・MAC", "エア・セーブMA", "エア・セーブMB", "エア・セーブMC", "エア・セーブME(カーテン)", "エア・セーブ"}
    opening_totals = {}
    curtain_names  = {}
    for it in (items or []):
        op = it.get("_open")
        if op is None:
            continue
        # 合計金額：メモ以外を加算（小計が無ければ 単価×数量）
        if it.get("種別") != "メモ":
            lt = it.get("小計")
            if lt in (None, ""):
                try:
                    lt = float(it.get("数量") or 0) * float(it.get("単価") or 0)
                except Exception:
                    lt = 0
            try:
                opening_totals[op] = (opening_totals.get(op, 0) or 0) + int(float(lt or 0))
            except Exception:
                pass
        # カーテン品名の捕捉（最初に出たカーテン種別の品名）
        if it.get("種別") in CURTAIN_KINDS and op not in curtain_names:
            curtain_names[op] = it.get("品名") or ""

    sum_row = 21
    for op in sorted(opening_totals.keys()):
        if sum_row > 44:
            break  # 44行まで
        ws0[f"A{sum_row}"].value = curtain_names.get(op, "")
        ws0[f"F{sum_row}"].value = 1
        ws0[f"G{sum_row}"].value = "式"
        ws0[f"H{sum_row}"].value = int(opening_totals.get(op, 0) or 0)
        sum_row += 1

    wb.save(out_path)


# =============================================
#  簡易テスト（Streamlit不要）
# =============================================
if __name__ == "__main__":
    from openpyxl import Workbook, load_workbook as _lw

    tpl = "_test_見積書テンプレ.xlsx"
    wb = Workbook(); wb.remove(wb.active)
    for name in ["見積書0"] + [f"見積書{i}" for i in range(1,6)]:
        ws = wb.create_sheet(title=name)
        if name == "見積書0":
            ws["J1"].value = "見積番号"; ws["J3"].value = "作成日"
            ws["A6"].value = "得意先"; ws["A7"].value = "支店 営業所"; ws["A8"].value = "担当者"; ws["B17"].value = "物件名"
        else:
            ws["A11"].value = "品名"; ws["F11"].value = "数量"; ws["G11"].value = "単位"; ws["H11"].value = "単価"
    wb.save(tpl)

    header = {
        "estimate_no":"NO-0001",
        "date":"2025-09-06",
        "customer_name":"テスト株式会社",
        "branch_name":"東京支店",
        "office_name":"新宿営業所",
        "person_name":"山田太郎",
        "project_name":"サンプル案件",
    }

    # --- テスト1：基本動作（定型句は末尾、空行、合計） ---
    items1 = [
        {"品名":"S・MACカーテンA","数量":1,"単位":"式","単価":1000,"小計":1000,"種別":"S・MAC","_open":1},
        {"品名":"部材X","数量":2,"単位":"式","単価":500,"小計":1000,"種別":"部材","_open":1},
        {"品名":"（定型文）","数量":"","単位":"","単価":"","小計":"","種別":"メモ","備考":"※ 定型句A","_open":1},
        {"品名":"S・MACカーテンB","数量":1,"単位":"式","単価":2000,"小計":2000,"種別":"S・MAC","_open":2},
    ]
    out1 = "_test_out1.xlsx"
    export_quotation_book_preserve(out1, header, items1, template_path=tpl)
    bk1 = _lw(out1)
    ws0 = bk1["見積書0"]
    assert ws0["J1"].value == header["estimate_no"] and ws0["B17"].value == header["project_name"]
    ws1 = bk1["見積書1"]
    assert ws1["A11"].value == "品名"  # 11行目は触らない
    assert ws1["A12"].value == "S・MACカーテンA"
    assert ws1["F12"].value == 1 and ws1["G12"].value == "式" and ws1["H12"].value == 1000
    assert ws1["A14"].value == "※ 定型句A"  # 定型句は間口末尾にA列のみ
    assert ws0["A21"].value == "S・MACカーテンA" and ws0["H21"].value == 2000

    # --- テスト2：開口をまたがない（残り行が不足なら次ページ、1枚目に開口2は出ない） ---
    items2 = []
    # 開口1: 30行
    for i in range(30):
        items2.append({"品名":f"S・MAC-{i}","数量":1,"単位":"式","単価":100,"小計":100,"種別":"S・MAC","_open":1})
    # 開口2: 10行（1枚目の残りに入らないので見積書2の12行目から始まる）
    for i in range(10):
        items2.append({"品名":f"部材-{i}","数量":1,"単位":"式","単価":50,"小計":50,"種別":"部材","_open":2})
    out2 = "_test_out2.xlsx"
    export_quotation_book_preserve(out2, header, items2, template_path=tpl)
    bk2 = _lw(out2)
    ws1b = bk2["見積書1"]; ws2b = bk2["見積書2"]
    # 1枚目のA列には『部材-0』が現れない（= 開口2は2枚目から）
    assert all((ws1b[f"A{i}"].value != "部材-0") for i in range(12, 45))
    # 2枚目の12行目は『部材-0』で開始
    assert ws2b["A12"].value == "部材-0"

    # --- テスト3：1開口が33行を超える場合はページをまたいでOK ---
    items3 = []
    for i in range(35):  # 35行（= 33超）
        items3.append({"品名":f"超長-{i}","数量":1,"単位":"式","単価":10,"小計":10,"種別":"S・MAC","_open":1})
    out3 = "_test_out3.xlsx"
    export_quotation_book_preserve(out3, header, items3, template_path=tpl)
    bk3 = _lw(out3)
    ws1c = bk3["見積書1"]; ws2c = bk3["見積書2"]
    assert ws1c["A44"].value is not None and ws2c["A12"].value is not None

    # --- テスト4：5ページ超でエラー ---
    try:
        items4 = []
        # 6開口 × 33行 で 6ページ必要 → エラー
        for op in range(1, 7):
            for i in range(33):
                items4.append({"品名":f"OP{op}-{i}","数量":1,"単位":"式","単価":1,"小計":1,"種別":"S・MAC","_open":op})
        export_quotation_book_preserve("_test_out4.xlsx", header, items4, template_path=tpl)
        raise AssertionError("5ページ超エラーが発生していません。")
    except ValueError as e:
        assert "5ページを超えました。生成不可" in str(e)

    print("[OK] すべての簡易テストを通過（33行超はページまたぎOK／5ページを超えたらエラー／開口単位でページをまたがない）。")
