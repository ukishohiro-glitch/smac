# -*- coding: utf-8 -*-
# main.py - Streamlit UI
# 要件:
# - ヘッダー入力（見積番号=3709-xxxxx固定、再生成不可）
# - マスタ検索→明細追加、手入力行も可
# - 間口ごとに入力（A=品名/F=数量/G=単位/H=単価）
# - 定型句は間口末尾のA列
# - 運賃・梱包は見積書0末尾（明細には載せない）
# - ページングルール：33行/頁、1間口≤33行は跨ぎ禁止、>33は跨ぎ可、5頁超はエラー
# - 見積書0: 21〜44行に間口合計を連続記載（空行なし）、直後に運賃・梱包
# - Excel保存はテンプレ保持（見積書0/1〜5）／保険として単一シート出力

import os, os.path as osp, re, secrets
from pathlib import Path
from datetime import date, datetime
from typing import List, Dict, Any

import streamlit as st
from openpyxl import load_workbook

try:
    APP_DIR = Path(__file__).parent
except NameError:
    APP_DIR = Path.cwd()
TEMPLATE_BOOK = APP_DIR / "お見積書（明細）.xlsx"
MASTER_BOOK   = APP_DIR / "master.xlsx"

from quote_logic import clean_openings, plan_paging, validate
from excel_export import export_quotation_book_preserve, export_detail_xlsx_preserve


# ============ 共通ユーティリティ ============
def _next_estimate_no() -> str:
    used = st.session_state.setdefault("used_estnos", set())
    for _ in range(2000):
        no = "3709-" + f"{secrets.randbelow(100000):05d}"
        if no not in used:
            used.add(no)
            return no
    return "3709-" + datetime.now().strftime("%H%M%S")[-5:]


def _load_master(path: Path) -> Dict[str, List[Dict[str, Any]]]:
    cat: Dict[str, List[Dict[str, Any]]] = {}
    if not path.exists(): return cat
    wb = load_workbook(str(path), data_only=True, read_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows: continue
        header = [str(x or "").strip() for x in rows[0]]
        idx = {h: i for i,h in enumerate(header)}
        def gi(*names):
            for nm in names:
                if nm in idx: return idx[nm]
            for nm in names:
                for h,i in idx.items():
                    if nm in h: return i
            return None
        c_name = gi("品名","商品名","名称","製品名","品番")
        if c_name is None: continue
        c_unit = gi("単位","Unit")
        c_price= gi("単価","価格","金額")
        c_model= gi("型番","型式","コード")
        items=[]
        for r in rows[1:]:
            if not r: continue
            name = str(r[c_name] if c_name<len(r) else "").strip()
            if not name: continue
            unit = str(r[c_unit] if c_unit is not None and c_unit<len(r) else "式").strip() or "式"
            rawp = r[c_price] if c_price is not None and c_price<len(r) else 0
            try: price = int(float(str(rawp).replace(",",""))); price = max(price,0)
            except: price = 0
            model = str(r[c_model] if c_model is not None and c_model<len(r) else "").strip()
            items.append({"品名":name,"単位":unit,"単価":price,"型番":model})
        if items: cat[ws.title]=items
    return cat


def _ensure_session():
    ss=st.session_state
    ss.setdefault("estimate_no", _next_estimate_no())
    ss.setdefault("file_title","見積")
    ss.setdefault("customer_name","")
    ss.setdefault("branch_name","")
    ss.setdefault("office_name","")
    ss.setdefault("rep_name","")
    ss.setdefault("project_name","")
    ss.setdefault("catalog", _load_master(MASTER_BOOK))
    ss.setdefault("search_query","")
    ss.setdefault("ship_option","（路線便時間指定不可）")
    ss.setdefault("ship_price",0)
    if "openings" not in ss:
        ss.openings=[{"name":"開口1","items":[{"品名":"","数量":1,"単位":"式","単価":0}],"teikeiku":[] }]


# ============ UI ============
def _header_block():
    st.set_page_config(layout="wide",page_title="お見積書作成")
    st.title("お見積書作成システム")
    c1,c2,c3=st.columns([1.2,0.9,1.0])
    with c1: st.text_input("見積番号",value=st.session_state.estimate_no,key="estimate_no",disabled=True)
    with c2: st.text_input("保存ファイル名",value=st.session_state.file_title,key="file_title")
    with c3: st.text_input("作成日",value=date.today().strftime("%Y/%m/%d"),key="created_disp",disabled=True)
    d1,d2,d3,d4=st.columns(4)
    with d1: st.text_input("得意先名",value=st.session_state.customer_name,key="customer_name")
    with d2: st.text_input("支店名",value=st.session_state.branch_name,key="branch_name")
    with d3: st.text_input("営業所名",value=st.session_state.office_name,key="office_name")
    with d4:
        st.text_input("担当者名(半角2〜4)",value=st.session_state.rep_name,key="rep_name")
        rep=st.session_state.get("rep_name") or ""
        if rep and not re.fullmatch(r"[A-Za-z0-9]{2,4}",rep):
            st.error("担当者名は半角英数2〜4文字で入力してください。")
    st.text_input("物件名",value=st.session_state.project_name,key="project_name")
    st.divider()


def _picker_block():
    st.markdown("### 部材マスタ追加（任意）")
    cat=st.session_state.catalog
    if not cat:
        st.info("master.xlsx が無いため手入力のみ。"); return
    cat_names=sorted(cat.keys())
    c=st.columns([2,2,3,1,1])
    with c[0]: csel=st.selectbox("カテゴリ",options=cat_names,key="cat_sel")
    items=cat.get(csel,[])
    with c[1]: st.text_input("検索",value=st.session_state.get("search_query",""),key="search_query_in")
    q=(st.session_state.get("search_query_in") or "").strip()
    filtered=[it for it in items if q.lower() in (it["品名"]+" "+(it.get("型番") or "")).lower()] if q else items
    names=[f'{it["品名"]}〔{it.get("型番","") or "-"} / {it["単価"]}円/{it["単位"]}〕' for it in filtered] or ["(該当なし)"]
    with c[2]: sel=st.selectbox("部材選択",options=names,key="part_sel")
    with c[3]: qty=st.number_input("数量",min_value=1,step=1,value=1,key="pick_qty")
    with c[4]: opi=st.number_input("間口No",min_value=1,step=1,value=1,key="pick_opening_idx")
    def _add():
        if not filtered: return
        i=names.index(sel) if sel in names else -1
        if i<0: return
        it=dict(filtered[i]); it["数量"]=qty
        while len(st.session_state.openings)<opi:
            st.session_state.openings.append({"name":f"開口{len(st.session_state.openings)+1}","items":[],"teikeiku":[]})
        st.session_state.openings[opi-1]["items"].append({"品名":it["品名"],"数量":it["数量"],"単位":it["単位"],"単価":it["単価"]})
    st.button("＋追加",on_click=_add,key="add_from_master")
    st.divider()


def _openings_block():
    st.markdown("### 間口明細")
    for oi,op in enumerate(st.session_state.openings):
        with st.expander(f"間口{oi+1}:{op.get('name')}",expanded=True):
            op["name"]=st.text_input("間口名",value=op.get("name",f"開口{oi+1}"),key=f"op{oi}_name")
            delrow=None
            for ri,row in enumerate(op["items"]):
                c=st.columns([5,1,1,2,0.9])
                with c[0]: row["品名"]=st.text_input("品名",value=row.get("品名",""),key=f"op{oi}_r{ri}_nm")
                with c[1]: row["数量"]=st.number_input("数量",min_value=0,step=1,value=int(row.get("数量",1)),key=f"op{oi}_r{ri}_qty")
                with c[2]: row["単位"]=st.text_input("単位",value=row.get("単位","式"),key=f"op{oi}_r{ri}_unit")
                with c[3]: row["単価"]=st.number_input("単価",min_value=0,step=100,value=int(row.get("単価",0)),key=f"op{oi}_r{ri}_pr")
                with c[4]:
                    if st.button("−",key=f"op{oi}_r{ri}_del"): delrow=ri
            if delrow is not None and 0<=delrow<len(op["items"]): del op["items"][delrow]
            if st.button("＋行追加",key=f"op{oi}_addrow"): op["items"].append({"品名":"","数量":1,"単位":"式","単価":0})
            tx="\n".join(op.get("teikeiku") or [])
            tx=st.text_area("定型句",value=tx,key=f"op{oi}_teikei")
            op["teikeiku"]=[s for s in tx.splitlines() if s.strip()]
            if st.button("×間口削除",key=f"op{oi}_del") and len(st.session_state.openings)>1:
                st.session_state.openings.pop(oi); st.experimental_rerun()
    if st.button("＋間口追加",key="add_opening"):
        st.session_state.openings.append({"name":f"開口{len(st.session_state.openings)+1}","items":[{"品名":"","数量":1,"単位":"式","単価":0}],"teikeiku":[]})
    st.divider()
    st.session_state.overall_items=clean_openings([{"name":op.get("name",""),"items":list(op.get("items") or []),"teikeiku":list(op.get("teikeiku") or [])} for op in st.session_state.openings])
    if st.session_state.overall_items:
        try:
            p=plan_paging(st.session_state.overall_items,rows_per_page=33)
            if p>5: st.warning(f"この入力は {p} ページ相当（上限5）。保存時にエラー。")
        except: pass


def _shipping_block():
    st.markdown("### 運賃・梱包（見積書0のみ）")
    a,b=st.columns(2)
    with a: st.radio("配送条件",options=["（路線便時間指定不可）","（現場搬入時間指定可）"],
                     index=0 if st.session_state.ship_option.startswith("（路線便") else 1,
                     key="ship_option",horizontal=True)
    with b: st.number_input("金額(円)",min_value=0,step=100,value=int(st.session_state.ship_price),key="ship_price")
    st.divider()


def _save_block():
    st.markdown("### 保存")
    if st.button("Excel保存",key="save_btn"):
        items=st.session_state.get("overall_items",[])
        if not items: st.error("明細なし"); st.stop()
        header={"見積番号":st.session_state.estimate_no,"作成日":date.today().strftime("%Y/%m/%d"),
                "得意先名":st.session_state.customer_name,"支店名":st.session_state.branch_name,
                "営業所名":st.session_state.office_name,"担当者名":st.session_state.rep_name,
                "物件名":st.session_state.project_name,
                "shipping":{"label":"運賃・梱包","option":st.session_state.ship_option,
                            "qty":1,"unit":"式","price":int(st.session_state.ship_price)}}
        try:
            validate(items,header,rows_per_page=33,max_pages=5)
            out=osp.join(os.getcwd(),f"{st.session_state.get('file_title','見積')}_お見積書（明細）.xlsx")
            tpl=str(TEMPLATE_BOOK)
            use_preserve=False
            try:
                wb=load_workbook(tpl); sn=set(wb.sheetnames)
                use_preserve=("見積書0" in sn) and all(f"見積書{i}" in sn for i in range(1,6))
            except: use_preserve=False
            if use_preserve:
                export_quotation_book_preserve(out,header,items,template_path=tpl,
                                               header_sheet="見積書0",detail_sheets=[f"見積書{i}" for i in range(1,6)],
                                               start_row=12,end_row=44)
            else:
                export_detail_xlsx_preserve(out,header,items,template_path=tpl,
                                            ws_name="お見積書（明細）",start_row=12,max_rows=33)
            st.success(f"保存完了: {out}")
        except ValueError as ve: st.error(str(ve))
        except Exception as e: st.error("Excel出力エラー"); st.exception(e)


def main():
    _ensure_session()
    _header_block(); _picker_block(); _openings_block(); _shipping_block(); _save_block()

if __name__=="__main__": main()
