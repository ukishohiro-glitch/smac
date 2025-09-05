# excel_export_com.py
# Windows + Excel 必須（pip install pywin32）
import os, shutil
from typing import List, Dict, Any, Optional

SHEETS_DETAIL = [f"見積書{i}" for i in range(1, 6)]  # 1..5
FIRST_ROW, LAST_ROW = 12, 44

def _fmt_dim(prefix: str, w, h) -> str:
    def _i(v):
        try:
            if v is None or v == "": return 0
            return int(float(v))
        except Exception:
            return 0
    return f"{prefix}W{_i(w):04d}×H{_i(h):04d}"

def _select_price_target(op: Dict[str, Any]) -> str:
    t = (op.get("price_target") or "").strip().lower()
    if t in ("air_save", "smac"): return t
    s = (op.get("curtain_subtype") or "").lower()
    if ("air" in s and "save" in s) or ("エア" in s and "セーブ" in s): return "air_save"
    if "smac" in s or "s・mac" in s or "ｓ・ｍａｃ" in s: return "smac"
    return "smac"

def _ensure_sheet(wb, name: str):
    """指定名のシートが無ければ追加して返す"""
    for ws in wb.Worksheets:
        if ws.Name == name:
            return ws
    ws_new = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws_new.Name = name
    return ws_new

def write_estimate_to_template(
    template_path: str,
    output_path: str,
    header: Dict[str, Any],
    openings: List[Dict[str, Any]],
    cover_total_cell: Optional[str] = None,
) -> None:
    """
    使い方（例）
      write_estimate_to_template(
          r"C:\\path\\to\\見積書テンプレ.xlsx",
          r"C:\\path\\to\\見積書_出力.xlsx",
          header, openings
      )
    """
    import win32com.client as win32

    # 1) 出力ファイルをテンプレから作成
    template_path = os.path.abspath(template_path)
    output_path   = os.path.abspath(output_path)
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"テンプレートが見つかりません: {template_path}")
    shutil.copyfile(template_path, output_path)

    # 2) Excel 起動
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(output_path)

        # ---------- 表紙（見積書0） ----------
        ws0 = _ensure_sheet(wb, "見積書0")
        ws0.Range("I1").Value = header.get("estimate_no", "")
        ws0.Range("I3").Value = header.get("date", "")
        ws0.Range("A6").Value = header.get("customer_name", "")
        ws0.Range("A7").Value = f"{header.get('branch_name','')} {header.get('office_name','')}".strip()
        ws0.Range("A8").Value = header.get("person_name", "")

        row = 21
        for op in openings:
            if row > 44: break
            ws0.Range(f"A{row}").Value = op.get("curtain_subtype", "")
            ws0.Range(f"F{row}").Value = "" if op.get("qty") is None else op.get("qty")
            ws0.Range(f"G{row}").Value = "" if op.get("unit") is None else op.get("unit")
            ws0.Range(f"H{row}").Value = "" if op.get("unit_price") is None else op.get("unit_price")
            row += 2
        if cover_total_cell:
            ws0.Range(cover_total_cell).Value = len(openings)

        # ---------- 明細（見積書1..5） ----------
        # 必要な明細シートを用意
        for name in SHEETS_DETAIL:
            _ensure_sheet(wb, name)

        sheet_idx = 0
        ws = wb.Worksheets(SHEETS_DETAIL[sheet_idx])
        row = FIRST_ROW

        def writeln(text: str, qty=None, unit=None, price=None):
            nonlocal row, sheet_idx, ws
            if row > LAST_ROW:
                sheet_idx += 1
                if sheet_idx >= len(SHEETS_DETAIL):
                    raise ValueError("明細ページが5枚を超えました。テンプレ増設が必要です。")
                ws = wb.Worksheets(SHEETS_DETAIL[sheet_idx])
                row = FIRST_ROW
            ws.Range(f"A{row}").Value = text
            # F/G/H は毎行確実に上書き
            ws.Range(f"F{row}").Value = "" if qty   is None else qty
            ws.Range(f"G{row}").Value = "" if unit  is None else unit
            ws.Range(f"H{row}").Value = "" if price is None else price
            row += 1

        for op in openings:
            sign = (op.get("sign") or "").strip()
            if sign: writeln(sign)

            subtype = (op.get("curtain_subtype") or "").strip()
            open_method = (op.get("open_method") or "").strip()
            title = f"{subtype} {open_method}".strip()
            if title: writeln(title)

            name = (op.get("product_name") or "").strip()
            if name: writeln(name)

            perf = (op.get("performance") or "").strip()
            if perf: writeln(perf)

            opening_line = _fmt_dim("間口寸法 ", op.get("opening_w"), op.get("opening_h"))
            curtain_line = _fmt_dim("カーテン寸法 ", op.get("curtain_w"), op.get("curtain_h"))
            qty, unit, price = op.get("qty"), op.get("unit"), op.get("unit_price")
            target = _select_price_target(op)

            if target == "air_save":
                # エア・セーブ：間口寸法に 金額
                writeln(opening_line, qty, unit, price)
            else:
                # S・MAC：間口寸法は見出し、カーテン寸法に 金額
                writeln(opening_line)
                writeln(curtain_line, qty, unit, price)

            # 任意部材
            for label, key in [
                ("カーテンレール", "rail"),
                ("取手付間仕切ポール", "pole_handle"),
                ("中間ポール", "middle_pole"),
                ("落し", "drop_bar"),
                ("その他", "others"),
                ("梱包・運賃", "packing_shipping"),
            ]:
                v = (op.get(key) or "").strip()
                if v: writeln(f"{label}：{v}")

            for p in (op.get("phrases") or []):
                s = (str(p) or "").strip()
                if s: writeln(s)

            writeln("")  # 区切り空行

        wb.Save()
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()
