# -*- coding: utf-8 -*-
# quote_logic.py - 見積ロジック層（入力クリーニング / ページ割 / バリデーション）
from typing import List, Dict, Any
from dataclasses import dataclass
import math
import re

Item    = Dict[str, Any]   # {"品名": str, "数量": int, "単位": str, "単価": int}
Opening = Dict[str, Any]   # {"name": str, "items": List[Item], "teikeiku": List[str]}
Header  = Dict[str, Any]   # {"見積番号","作成日","得意先名","支店名","営業所名","担当者名","物件名","shipping":{...}}

ROWS_PER_PAGE  = 33   # 明細：12〜44行（11行目は触らない）
SHEET0_START   = 21
SHEET0_END     = 44
SHEET0_CAP     = SHEET0_END - SHEET0_START + 1   # 24 行

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
    """
    - 空の明細行を除去 / 数量・単価は非負整数化
    - 入力順は保持
    """
    cleaned: List[Opening] = []
    for op in openings or []:
        name = str(op.get("name") or "").strip()
        items: List[Item] = []
        for it in op.get("items") or []:
            nm = str(it.get("品名") or "").strip()
            if not nm: 
                continue
            items.append({
                "品名": nm,
                "数量": _to_int_nonneg(it.get("数量", 0)),
                "単位": str(it.get("単位") or "式").strip() or "式",
                "単価": _to_int_nonneg(it.get("単価", 0)),
            })
        teikeiku = [str(s).strip() for s in (op.get("teikeiku") or []) if str(s).strip()]
        if items or teikeiku or name:
            cleaned.append({"name": name, "items": items, "teikeiku": teikeiku})
    return cleaned

def opening_total(items: List[Item]) -> int:
    tot = 0
    for it in items or []:
        tot += _to_int_nonneg(it.get("数量", 0)) * _to_int_nonneg(it.get("単価", 0))
    return tot

def first_curtain_name(items: List[Item]) -> str:
    """間口内で最初に登場するカーテン種別の品名（S・MAC→エア・セーブ各型→ME→「カーテン」→先頭品名）。"""
    if not items: return ""
    keywords = ["S・MAC", "エア・セーブ", "ME", "カーテン"]
    for it in items:
        nm = str(it.get("品名") or it.get("name") or it.get("product") or "")
        for kw in keywords:
            if kw in nm: return nm
    return str(items[0].get("品名") or items[0].get("name") or items[0].get("product") or "")

def is_shipping_required(flat_items: List[Item]) -> bool:
    """S・MAC または エア・セーブ MA 型を含む → 梱包・運賃 必須"""
    for it in flat_items or []:
        name = str(it.get("品名") or it.get("name") or it.get("product") or "")
        if "S・MAC" in name:
            return True
        if "エア・セーブ" in name and re.search(r"\bMA\b|MA型", name):
            return True
    return False

@dataclass(frozen=True)
class PageCell:
    page_index: int   # 0-based（見積書1=0 ...）
    row_index: int    # 実シートの行番号（12〜44）
    kind: str         # "item" or "teikei"
    payload: Any
    opening_idx: int

def _need_rows(op: Opening) -> int:
    return len(op.get("items") or []) + len(op.get("teikeiku") or [])

def plan_paging(openings: List[Opening], rows_per_page: int = ROWS_PER_PAGE) -> int:
    """概算ページ数。1間口≤33は同一ページ内完結、>33は分割可。"""
    if not openings: return 0
    pages = 1
    remain = rows_per_page
    for op in openings:
        need = _need_rows(op)
        if need <= rows_per_page:
            if need > remain:
                pages += 1
                remain = rows_per_page
            remain -= need
            if remain == 0:
                pages += 1
                remain = rows_per_page
        else:
            rest = need
            while rest > 0:
                fit = min(remain, rest)
                rest -= fit
                remain -= fit
                if rest > 0:
                    pages += 1
                    remain = rows_per_page
                elif remain == 0:
                    pages += 1
                    remain = rows_per_page
    return max(1, pages)

def sheet0_required_rows(openings: List[Opening], header: Header) -> int:
    """見積書0の必要行数＝間口数 + (運賃・梱包>0なら+1)"""
    n = len(openings or [])
    ship = header.get("shipping") or {}
    price = _to_int_nonneg(ship.get("price", 0))
    return n + (1 if price > 0 else 0)

def assert_sheet0_fits(openings: List[Opening], header: Header) -> None:
    if sheet0_required_rows(openings, header) > SHEET0_CAP:
        raise ValueError("見積書0の行数が 21〜44 を超えました。")

def validate(openings_or_items: List[Opening], header: Header, *, rows_per_page: int = ROWS_PER_PAGE, max_pages: int = 5) -> None:
    """
    openings_or_items は：
      - 既に {"name","items","teikeiku"} 構造 → そのまま
      - 旧構造（itemsの配列だけ）→ 1つの間口と見做す
    """
    # 形式を正規化
    if openings_or_items and isinstance(openings_or_items[0], dict) and "items" not in openings_or_items[0]:
        openings = [{"name":"", "items": openings_or_items, "teikeiku": []}]
    else:
        openings = list(openings_or_items)

    openings = clean_openings(openings)

    # 梱包必須
    flat: List[Item] = []
    for op in openings:
        flat.extend(op.get("items") or [])
    if is_shipping_required(flat):
        price = _to_int_nonneg((header.get("shipping") or {}).get("price", 0))
        if price <= 0:
            raise ValueError("S・MAC または エア・セーブ MA 型を含むため、運賃・梱包の金額入力が必須です。")

    # 見積書0 行数
    assert_sheet0_fits(openings, header)

    # ページ数
    pages = plan_paging(openings, rows_per_page=rows_per_page)
    if pages > max_pages:
        raise ValueError("5ページを超えました。生成不可")
