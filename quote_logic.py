# -*- coding: utf-8 -*-
# quote_logic.py — 見積ロジック
# - 入力クリーニング
# - ページ割（33行/頁、1間口<=33は跨ぎ禁止、>33は跨ぎ可）
# - 見積書0の行数（21〜44=24行）制約
# - 梱包必須: S・MAC / エア・セーブMA（MB/MC/MEは任意）

from typing import List, Dict, Any
import math, re

Item    = Dict[str, Any]
Opening = Dict[str, Any]
Header  = Dict[str, Any]

ROWS_PER_PAGE = 33
SHEET0_CAP = 24  # 21〜44

def _to_int_nonneg(x: Any) -> int:
    try:
        if x is None: return 0
        if isinstance(x, int): return x if x >= 0 else 0
        if isinstance(x, float):
            if math.isnan(x) or math.isinf(x): return 0
            return int(x) if x >= 0 else 0
        s = str(x).strip().replace(",", "")
        if s == "": return 0
        v = int(float(s))
        return v if v >= 0 else 0
    except Exception:
        return 0

def clean_openings(openings: List[Opening]) -> List[Opening]:
    cleaned: List[Opening] = []
    for op in openings or []:
        name = str(op.get("name") or "").strip()
        items: List[Item] = []
        for it in op.get("items") or []:
            nm = str(it.get("品名") or "").strip()
            if not nm: continue
            items.append({
                "品名": nm,
                "数量": _to_int_nonneg(it.get("数量", 0)),
                "単位": str(it.get("単位") or "式").strip() or "式",
                "単価": _to_int_nonneg(it.get("単価", 0)),
            })
        teikei = [str(s).strip() for s in (op.get("teikeiku") or []) if str(s).strip()]
        if items or teikei or name:
            cleaned.append({"name": name, "items": items, "teikeiku": teikei})
    return cleaned

def is_shipping_required(flat_items: List[Item]) -> bool:
    for it in flat_items or []:
        nm = str(it.get("品名") or "")
        if "S・MAC" in nm:
            return True
        if "エア・セーブ" in nm and re.search(r"\bMA\b|MA型", nm):
            return True
    return False

def plan_paging(openings: List[Opening], rows_per_page: int = ROWS_PER_PAGE) -> int:
    if not openings: return 0
    pages, remain = 1, rows_per_page
    for op in openings:
        need = len(op.get("items") or []) + len(op.get("teikeiku") or [])
        if need <= rows_per_page:
            if need > remain:
                pages += 1; remain = rows_per_page
            remain -= need
            if remain == 0:
                pages += 1; remain = rows_per_page
        else:
            rest = need
            while rest > 0:
                fit = min(remain, rest)
                rest -= fit; remain -= fit
                if rest > 0:
                    pages += 1; remain = rows_per_page
                elif remain == 0:
                    pages += 1; remain = rows_per_page
    return max(1, pages)

def sheet0_required_rows(openings: List[Opening], header: Header) -> int:
    n = len(openings or [])
    ship = header.get("shipping") or {}
    price = _to_int_nonneg(ship.get("price", 0))
    return n + (1 if price > 0 else 0)

def validate(openings_or_items: List[Opening], header: Header, *,
             rows_per_page: int = ROWS_PER_PAGE, max_pages: int = 5) -> None:
    openings = list(openings_or_items or [])
    openings = clean_openings(openings)

    # 梱包必須
    flat: List[Item] = []
    for op in openings: flat.extend(op.get("items") or [])
    if is_shipping_required(flat):
        price = _to_int_nonneg((header.get("shipping") or {}).get("price", 0))
        if price <= 0:
            raise ValueError("S・MAC または エア・セーブ MA 型を含むため、運賃・梱包の金額入力が必須です。")

    # 見積書0の行数
    if sheet0_required_rows(openings, header) > SHEET0_CAP:
        raise ValueError("見積書0（21〜44行）に収まりません。")

    # 明細ページ数
    pages = plan_paging(openings, rows_per_page=rows_per_page)
    if pages > max_pages:
        raise ValueError("5ページを超えました。生成不可")
