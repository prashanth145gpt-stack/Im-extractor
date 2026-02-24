from fastapi import FastAPI, UploadFile, File, Request,Form, Response
import tempfile, os, json, re, time, subprocess
import pdfplumber, requests, traceback
from bs4 import BeautifulSoup
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse, PlainTextResponse
from typing import Any, Dict, Union, List
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware
from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from io import BytesIO
from typing import Optional
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.document import Document as DocxDocument
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, NoSuchElementException
# from selenium.webdriver.edge.service import Service
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.edge.service import Service as EdgeService
# from selenium.webdriver.edge.options import Options
import os, time, json, subprocess, tempfile, re
from lxml import html
import time, requests, urllib3 
import xml.etree.ElementTree as ET
from urllib3.util.retry import Retry
from requests.adapters import HTTPAdapter
from pydantic import BaseModel, Field, field_validator
import time, pymupdf

FONT_NAME = "Rupee Foradian"
max_retry = 2
page_load_timeout = 30

Number = Union[int,float]
PROXY = "http://172.16.180.43:80"

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins =[
        "*"
    ],
    allow_credentials = True,
    allow_methods = ["*"],
    allow_headers = "*",
)

class HtmlToDocxPayload(BaseModel):
    html: str
    file_name: Optional[str] = "Output.docx"

class FillHtmlPayload(BaseModel):
    data:List[Dict[str,Any]]
    template_html:str

# class FillHtmlRequest(BaseModel):
#     html:str
#     extracted_json : List[Dict[str,Any]]    

def get_cell_text(cell):
    t = cell.text
    if t is None: return None
    t = t.strip()
    return t if t else None

def is_float(val):
    try:
        float(val)
        return True
    except ValueError:
        return False


def to_crores(val):
    if val is None:
        return "Couldn't be Captured"
    if isinstance(val, str):
        v = val.strip()
        if v == "" or v == "NA":
            return "Couldn't be Captured"
        # low = v.lower().replace(",", "")
        # nums = re.search(r"([0-9]+(?:\.[0-9]+)?)\s*lakhs?", low)
        # if nums:
        #     num = float(nums.group(1))
        #     return f"{round(num/100,2)}"
    if isinstance(val,(int,float)):
        return f"{val}"
    return str(val)

def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    for edge in ("top", "left", "bottom", "right"):
        tag = f"w:{edge}"
        element = OxmlElement(tag)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "6")   # thin line
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "000000")  # black (change D9D9D9 for light gray)
        tcPr.append(element)


def cell_merged(cell):
    return getattr(cell._tc, "gridSpan", 1) > 1


def add_text_with_missing_highlight(p, text, size, bold=False):
    """
    Adds text to paragraph and highlights only
    'Couldn't be captured' portion in red + underline.
    """

    if not text:
        return

    lower_txt = text.lower()
    keyword = "couldn't be captured"

    if keyword in lower_txt:
        start = lower_txt.index(keyword)
        end = start + len(keyword)

        before = text[:start]
        missing = text[start:end]
        after = text[end:]

        # normal before
        if before:
            r1 = p.add_run(before)
            format_run(r1, size, bold)

        # red missing part
        r2 = p.add_run(missing)
        format_run(r2, size, bold)
        r2.font.underline = True
        r2.font.color.rgb = RGBColor(218,65,8)

        # normal after
        if after:
            r3 = p.add_run(after)
            format_run(r3, size, bold)

    else:
        r = p.add_run(text)
        format_run(r, size, bold)



def format_run(run,size,bold:bool = False,underline:bool=False):
    run.bold = bold
    run.underline = underline
    run.font.name = FONT_NAME
    run.font.size = size
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)

    if run.text and "couldn't be captured" in run.text.lower():
        run.font.color.rgb = RGBColor(218,65,8)


def add_text(container, text: str, size: int, bold: bool = False, underline:bool = False):
    """
    Add text with exact font, size and NO Word style
    """
    if isinstance(container, DocxDocument):
        p = container.add_paragraph()
    else:
        p = container

    p.style = None  # 🔥 no Word style

    run = p.add_run(text)
    format_run(run,size,bold,underline)

    return p


# def add_text(container, text: str, size: int, bold: bool = False):
#     """
#     Add text with exact font, size and NO Word style
#     """
#     if isinstance(container, DocxDocument):
#         p = container.add_paragraph()
#     else:
#         p = container

#     p.style = None  # 🔥 no Word style

#     run = p.add_run(text)
#     run.bold = bold
#     run.font.name = FONT_NAME
#     run.font.size = Pt(size)
#     run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)

#     return p


def is_fully_merged_row(row):
    spans = [getattr(c._tc, "gridSpan", 1) for c in row.cells]
    return max(spans) > 1 and len(set(spans)) == 1


def is_srno(x):
    if x is None: return False
    x2 = x.strip().lower().replace('.', '').replace(' ', '').replace('\n','')
    if x2.isdigit(): return True
    return x2 in ["srno", "sr","sno"]


def clean_keys(obj):
    if isinstance(obj,dict):
        new = {}
        for k,v in obj.items():
            k = k.replace("\n"," ").strip()
            k = "_".join(k.split())
            new[k] = clean_keys(v)
        return new
    if isinstance(obj,list):
        return[clean_keys(i) for i in obj]
    if isinstance(obj,str):
        obj = obj.replace("\n"," ").strip()
        # obj = obj.replace("/"," ").strip()
        obj = "_".join(obj.split())
        return obj
    else:
        obj = obj.replace(" ","_")
    return obj


def extract_docx_json(path):
    doc = Document(path)
    results = []

    #  multi-level headers → single flat names
    # headers: list of header rows
    # idx: current row index
    # prefix: already built string
    # CARTESIAN PRODUCT
    def combine_headers(headers, idx=0, prefix=""):
        if idx >= len(headers): return [prefix] # "sab rows process ho gayi → final name ready"
        row = headers[idx]
        result = []
        for col in row:
            col_clean = col.strip() if col else ""
            new_prefix = f"{prefix}_{col_clean}" if prefix else col_clean
            result.extend(combine_headers(headers, idx+1, new_prefix))
        return result

    for el in doc.element.body:
        if el.tag.endswith('tbl'):
            tbl = None
            for t in doc.tables:
                if t._element == el:
                    tbl = t
                    break
            if tbl is None: continue
            rows_text = [[get_cell_text(c) for c in r.cells] for r in tbl.rows]
            if not rows_text: continue

            has_partial = any(
                any(v is None or v.strip() == "" for v in r) and
                any(v is not None and v.strip() != "" for v in r)
                for r in rows_text
            )

            cleaned_rows = []
            for r in rows_text:
                if all(v and v.strip() for v in r):
                    if has_partial:
                        cleaned_rows.append(r[1:])
                    else:
                        cleaned_rows.append(r)
                else:
                    cleaned_rows.append(r)
            rows_text = cleaned_rows

            first_cell = rows_text[0][0] if rows_text[0] else None
            ignore_first_col = is_srno(first_cell)

            processed_rows = []
            first_none_idx = None
            # Identifying the first empty row, skipping headers
            for idx, r in enumerate(rows_text):
                if any(v is None or v.strip() == "" for v in r):
                    first_none_idx = idx
                    break
            
            # Skipping merged rows, keeping only first one (even if merged)
            for idx, (r, orig_row) in enumerate(zip(rows_text, tbl.rows)):
                merged_full = is_fully_merged_row(orig_row)
                merged_any = any(cell_merged(c) for c in orig_row.cells)
                if merged_full or merged_any: # Keeping only first row header even if merged
                    if idx != 0:
                        continue
                processed_rows.append(r)

            if first_none_idx is not None:
                filtered_rows = []
                for idx, r in enumerate(processed_rows):
                    if first_none_idx < idx:
                        if all(v and v.strip() for v in r):
                            continue
                    filtered_rows.append(r)
                processed_rows = filtered_rows
            
            adjusted_rows = []
            for r in processed_rows:
                fully_filled = all(v and v.strip() for v in r) 
                if fully_filled and ignore_first_col and len(r) > 1: # First row header, and removed Sr.No
                    r = r[1:]
                adjusted_rows.append(r)
            processed_rows = adjusted_rows
            column_hierarchy = []
            row_headers = []

            for idx, r in enumerate(processed_rows):
                fully_filled = all(v and v.strip() for v in r)
                if fully_filled and (first_none_idx is None or idx < first_none_idx):
                    column_hierarchy.append(r)
                else:
                    cleaned = [v for v in r[:2] if v and v.strip()]
                    if cleaned:
                        if len(cleaned) == 1:
                            row_headers.append(cleaned[0])
                        else:
                            non_digits = [x for x in cleaned if not x.isdigit()]
                            if non_digits:
                                row_headers.append(non_digits[0])
                            else:
                                row_headers.append(cleaned[0])
        
            row_headers = clean_keys(row_headers)
            column_hierarchy = clean_keys(column_hierarchy)
            
            main_row_col_idx = None
            if row_headers and column_hierarchy:
                for r in processed_rows:
                    for i, v in enumerate(r):
                        if v and v.strip() and clean_keys(v) in row_headers: #type: ignore
                            main_row_col_idx = i
                            break
                    if main_row_col_idx is not None:
                        break
                if main_row_col_idx is not None:
                    column_hierarchy = [r[main_row_col_idx:] for r in column_hierarchy]

            flat_col_headers = []
            for r in column_hierarchy:
                for c in r:
                    if c not in flat_col_headers:
                        flat_col_headers.append(c)

            placeholders = []
            if column_hierarchy and row_headers:
                col_paths = combine_headers(column_hierarchy)
                for cp in row_headers:
                    for rh in col_paths:
                        placeholders.append(cp + "_" + rh)
            elif column_hierarchy:
                col_paths = combine_headers(column_hierarchy)
                for cp in flat_col_headers:
                    placeholders.append(cp)
            elif row_headers:
                for rh in row_headers:
                    placeholders.append(rh)

            placeholders = list(dict.fromkeys(placeholders))

            placeholders = list(dict.fromkeys(placeholders))

            results.append({
                "placeholders": {th: "" for th in placeholders}
            })

        elif el.tag.endswith('p'):
            p_text = ''.join([node.text for node in el if node.text])
            matches = re.findall(r'\[(.*?)\]', p_text)
            for m in matches:
                results.append({m.strip().replace(" ","_"): ""})

    return results


def normalize_key(k):
    k = k.strip()
    k = k.lower()
    k = re.sub(r'\s+', '_', k)
    k = re.sub(r'[^a-z0-9_]', '_', k)
    k = re.sub(r'_+', '_', k)
    k = k.strip('_')
    return k

def to_int(x: Any, field: str) -> int:
    if x is None:
        return 0    #return 0 instead of error
    if isinstance(x, int):
        return x
    if isinstance(x, float):
        if not x.is_integer():
            raise ValueError(f"{field} must be integer-like, got {x}")
        return int(x)
    if isinstance(x, str):
        s = x.strip().upper()
        if s in {"", "NIL", "-", "NA", "N/A","Cannot be extracted"}:
            return 0
        s = s.replace(",", "").replace(" ", "")
        num = re.search(r'\d+(?:\.\d+)?',s)
        if num:
            return int(float(num.group()))
    return 0    #return 0 instead of error.


def to_float(x: Any, field: str) -> float:
    if x is None:
        return 0   #return zero checker instead of error
    if isinstance(x, (int, float)):
        return float(x)
    if isinstance(x, str):
        s = x.strip().upper()
        if s in {"", "NIL", "-", "NA", "N/A","Cannot be extracted"}:
            return 0.0
        s = s.replace(",", "").replace(" ", "")
        num = re.search(r'\d+(?:\.\d+)?',s)
        if num:
            return float(num.group())
    return 0   #return zero instead of error


 
def lakh_to_rupees(x_lakh: Number) -> float:
    return float(x_lakh) * 100_000.0

def crores_to_lakhs(x_crores: Number) -> float:
    return float(x_crores) * 100
 
def rupees_to_crore(x_rupees: Number) -> float:
    return float(x_rupees) / 10_000_000.0
 
class SimulationRequest(BaseModel):
    pre_issue_shares: Any = Field(..., alias="Share_Holding_Pre_Issue")
    fresh_issue_shares: Any = Field(..., alias="Number_of_shares")
    ofs_promoter_shares: Any = Field(..., alias="Offer_for_sale_size")
    promoter_pre_issue_shares: Any = Field(..., alias="Subtotal_(A)_Shareholding_Pattern_Pre-IPO_Number_of_Shares")
    public_pre_issue_shares : Any = Field(..., alias="Subtotal_(B)_Shareholding_Pattern_Pre-IPO_Number_of_Shares")
    pre_issue_networth_lakh: Any = Field(..., alias="Net_Worth_31_Mar_2024")
    issue_price1: Any = Field(..., alias="issue_price1")
    issue_price3: Any = Field(..., alias="issue_price3")
    pat_lakh_25: Any = Field(..., alias="PAT_(`_in_lakh)_Mar-25")
    pat_lakh_24: Any = Field(..., alias="PAT_(`_in_lakh)_Mar-24")
    pat_lakh_23: Any = Field(..., alias="PAT_(`_in_lakh)_Mar-23")
    pat_backup_25 : Any = Field(..., alias="Profit_After_Tax_31_March_2025")
    pat_backup_24 : Any = Field(..., alias="Profit_After_Tax_31_Mar_2024")
    pat_backup_23 : Any = Field(..., alias="Profit_After_Tax_31_Mar_2023")
    ronw : Any = Field(..., alias="1._RONW(%)")
    bvp_issue : Any = Field(..., alias="1._B.V_Pre-Issue")
    pat_p1 : Any = Field(..., alias="Peer_1._PAT")
    eps_p1 : Any = Field(..., alias="Peer_1._EPS")
    ronw_p1 : Any = Field(..., alias="Peer_1._RONW(%)")
    pe_p1 : Any = Field(..., alias="Peer_1._P/E")
    bvp_p1 : Any = Field(..., alias="Peer_1._B.V_Pre-Issue")
    pat_p2 : Any = Field(..., alias="Peer_2._PAT")
    eps_p2 : Any = Field(..., alias="Peer_2._EPS")
    ronw_p2 : Any = Field(..., alias="Peer_2._RONW(%)")
    pe_p2 : Any = Field(..., alias="Peer_2._P/E")
    bvp_p2 : Any = Field(..., alias="Peer_2._B.V_Pre-Issue")
 
    issue_expenses_lakh: int = 0
    weights: Optional[Dict[str, int]] = Field(None)
 
    # @field_validator("pat_lakh")
    # @classmethod
 
    # def validate_pat(cls, v:Dict[str, Any]):
    #     if not isinstance(v, dict) or not v:
    #         raise ValueError("pat_lakh must be a non-empty dict")
       
    #     for k, val in v.items():
    #         try:
    #             float(str(val).replace(",", "").strip())
    #         except Exception:
    #             raise ValueError(f"pat_lakh[{k}] must be numeric, got: {val}")
    #     return v
 
# Core Formula
def compute_issue_simulation(axs, payload: SimulationRequest) -> Dict[str, Any]:
    pre_issue_shares = to_int(payload.pre_issue_shares, "Share_Holding_Pre_Issue")
    fresh_issue_shares = to_int(payload.fresh_issue_shares, "Number_of_shares")
    ofs_promoter_shares = to_int(payload.ofs_promoter_shares, "Offer_for_sale_size")
    promoter_pre_issue_shares = to_int(payload.promoter_pre_issue_shares, "Subtotal_(A)_Shareholding_Pattern_Pre-IPO_Number_of_Shares")
    public_pre_issue_shares = to_int(payload.public_pre_issue_shares,"Subtotal_(B)_Shareholding_Pattern_Pre-IPO_Number_of_Shares")
    pre_issue_networth_lakh = crores_to_lakhs(to_float(payload.pre_issue_networth_lakh, "Net_Worth_31_Mar_2024"))
    issue_price1 = to_float(payload.issue_price1, "issue_price1")
    issue_price2 = to_float(payload.issue_price3, "issue_price3")
    issue_expenses_lakh = 0
    pat_lakh_25 = to_float(payload.pat_lakh_25,"PAT_(`_in_lakh)_Mar-25")
    pat_lakh_24 = to_float(payload.pat_lakh_24,"PAT_(`_in_lakh)_Mar-24")
    pat_lakh_23 = to_float(payload.pat_lakh_23,"PAT_(`_in_lakh)_Mar-23")
    pat_backup_23 = to_float(payload.pat_backup_23,"Profit_After_Tax_31_Mar_2023")
    pat_backup_24 = to_float(payload.pat_backup_24,"Profit_After_Tax_31_Mar_2024")
    pat_backup_25 = to_float(payload.pat_backup_25,"Profit_After_Tax_31_March_2025")
    ronw = to_float(payload.ronw,"1._RONW(%)")
    bvp_issue = to_float(payload.bvp_issue, "1._B.V_Pre-Issue")
    pat_p1 = to_float(payload.pat_p1,"Peer_1._PAT")
    eps_p1 = to_float(payload.eps_p1,"Peer_1._EPS")
    ronw_p1 = to_float(payload.ronw_p1,"Peer_1._RONW(%)")
    pe_p1 = to_float(payload.pe_p1,"Peer_1._P/E")
    bvp_p1 = to_float(payload.bvp_p1,"Peer_1._B.V_Pre-Issue")
    pat_p2 = to_float(payload.pat_p2,"Peer_2._PAT")
    eps_p2 = to_float(payload.eps_p2,"Peer_2._EPS")
    ronw_p2 = to_float(payload.ronw_p2,"Peer_2._RONW(%)")
    pe_p2 = to_float(payload.pe_p2,"Peer_2._P/E")
    bvp_p2 = to_float(payload.bvp_p2,"Peer_2._B.V_Pre-Issue")
    
    total_pre_issue_shares = promoter_pre_issue_shares+public_pre_issue_shares
    
    if not pat_lakh_23:
        if pat_backup_23:
            pat_lakh_23= crores_to_lakhs(pat_backup_23)
    if not pat_lakh_24:
        if pat_backup_24:
            pat_lakh_24= crores_to_lakhs(pat_backup_24)
    if not pat_lakh_25:
        if pat_backup_25:
            pat_lakh_25= crores_to_lakhs(pat_backup_25)
            
    pat_latest = pat_lakh_25/100
    
    pat_lakh = {"Mar25":pat_lakh_25,"Mar24":pat_lakh_24,"Mar23":pat_lakh_23}
 
    weights = {"Mar25": 3, "Mar24": 2, "Mar23": 1}
    
    # issue shares
    total_issue_shares = fresh_issue_shares + ofs_promoter_shares
    post_issue_shares = pre_issue_shares + total_issue_shares
    results = []
    # promoter post-issue
    promoter_post_issue_shares = promoter_pre_issue_shares - ofs_promoter_shares
    promoter_post_issue_pct = (promoter_post_issue_shares / post_issue_shares)*100 if post_issue_shares else None
    public_post_issue_shares = post_issue_shares - promoter_post_issue_shares
    public_post_issue_pct = (public_post_issue_shares/post_issue_shares)*100 if post_issue_shares else None
    aif_post_issue_shares = public_post_issue_shares - fresh_issue_shares
    aif_post_issue_pct = (aif_post_issue_shares/post_issue_shares)*100 if post_issue_shares else None
    ipo_shareholders_pct = (fresh_issue_shares/post_issue_shares)*100 if post_issue_shares else None
    
    net_worth_1 = float((pat_p1*100)/ronw_p1) if ronw_p1 else None
    net_worth_2 = float((pat_p2*100)/ronw_p2) if ronw_p2 else None
    issue_price_p1 = float(pe_p1*eps_p1)
    issue_price_p2 = float(pe_p2*eps_p2)
    pbv_p1 = float(issue_price_p1/bvp_p1) if bvp_p1 else None
    pbv_p2 = float(issue_price_p2/bvp_p2) if bvp_p2 else None
    
    eps = {}
    eps_after_issue = {}
    for period, pat in pat_lakh.items():
        pat_rupees = lakh_to_rupees(pat)
        eps[period] = pat_rupees / pre_issue_shares if pre_issue_shares else None
        eps_after_issue[period] = pat_rupees / post_issue_shares if post_issue_shares else None
    
                # choose latest period
    latest_key = "Mar25" if "Mar25" in eps else next(iter(eps.keys()))
    eps_latest = eps.get(latest_key)
    eps_post_latest = eps_after_issue.get(latest_key)

                # weighted average EPS
    num = 0.0
    den = 0.0
    for period, e in eps.items():
        w = weights.get(period, 0)
        if e is not None and w > 0:
            num += e * w
            den += w
    wtd_avg_eps = num / den if den else None
    
    issue_price_avg = float((issue_price1+issue_price2)/2)

    pe_latest = issue_price2 / eps_latest if eps_latest else None
    pe_post_issue_latest = pe_latest / eps_post_latest if eps_post_latest else None
        
    net_worth_latest = (float((pat_lakh_25*100)/ronw) if (pat_lakh_25 and ronw) else None)

    pbv_self = float(issue_price_avg/bvp_issue) if bvp_issue else None   
    
    if axs:
        issue_prices = [issue_price1,float((issue_price1+issue_price2)/2),issue_price2]
        # issue size & valuation (₹ Cr)
        for i,issue_price in enumerate(issue_prices):
            fresh_proceeds_rupees = fresh_issue_shares * issue_price
            ofs_value_rupees = ofs_promoter_shares * issue_price
            total_issue_value_rupees = total_issue_shares * issue_price
            post_issue_valuation_rupees = post_issue_shares * issue_price
    
            fresh_issue_size_cr = rupees_to_crore(fresh_proceeds_rupees)
            ofs_issue_size_cr = rupees_to_crore(ofs_value_rupees)
            total_issue_size_cr = rupees_to_crore(total_issue_value_rupees)
            post_issue_valuation_cr = rupees_to_crore(post_issue_valuation_rupees)
            
            pe_ratio = float(issue_price/eps_latest) if eps_latest else None
            pe_post_ratio = float(pe_ratio/eps_post_latest) if eps_post_latest else None
            pe_weighted = float(issue_price / wtd_avg_eps) if wtd_avg_eps else None
    
        # EPS (pre and post)
                
            # PE ratios

            # BV & PBV
            pre_issue_networth_rupees = lakh_to_rupees(pre_issue_networth_lakh)
            bv_pre = pre_issue_networth_rupees / pre_issue_shares if pre_issue_shares else None
            pbv_pre = issue_price / bv_pre if bv_pre else None
        
            # Post-issue net worth model
            post_issue_networth_lakh = pre_issue_networth_lakh + (fresh_proceeds_rupees / 100_000.0) - issue_expenses_lakh
            post_issue_networth_rupees = lakh_to_rupees(post_issue_networth_lakh)
        
            bv_post = post_issue_networth_rupees / post_issue_shares if post_issue_shares else None
            pbv_post = issue_price / bv_post if bv_post else None
            
            results.append({
                "placeholders": {
                    "Share_Holding_Pre_Issue": pre_issue_shares,
                    "PAT_(`_in_lakh)_Mar-25" : pat_lakh_25,
                    "PAT_(`_in_lakh)_Mar-24" : pat_lakh_24,
                    "PAT_(`_in_lakh)_Mar-23" : pat_lakh_23,
                    "1._PAT" : pat_latest,
                    "1._Net_worth" : net_worth_latest,
                    "1._PBV" : pbv_self,
                    "issue_size": fresh_issue_shares,
                    "ofs_promoter_shares": ofs_promoter_shares,
                    "Subtotal_(A)_Shareholding_Pattern_Pre-IPO_Number_of_Shares": int(promoter_pre_issue_shares),
                    f"Price_(Rs)_Mar-{25-i}": issue_price,
                    "pre_issue_networth_lakh": pre_issue_networth_lakh,
                    "issue_expenses_lakh": issue_expenses_lakh,
                    "pat_lakh": pat_lakh,
                    "1._P/E" : pe_latest,
                    "1._EPS" : eps_latest,
                    "Peer_1._Net_worth": net_worth_1,
                    "Peer_1._Issue_Price": issue_price_p1,
                    "Peer_1._PBV": pbv_p1,
                    "Peer_2._Net_worth": net_worth_2,
                    "Peer_2._Issue_Price": issue_price_p2,
                    "Peer_2._PBV": pbv_p2,
                    "1._Issue_Price" : issue_price_avg,
                    f"Weight_Mar-{25-i}": weights[f"Mar{25-i}"],
                    f"EPS_Mar-{25-i}" : eps[f"Mar{25-i}"],
                    f"EPS_after_adjusting_Bonus_Mar-{25-i}" : eps[f"Mar{25-i}"],
                    f"EPS_after_Proposed_Issue_Mar-{25-i}" : eps_after_issue[f"Mar{25-i}"],
                    "weighted_avg_eps" : wtd_avg_eps,
                    f"EPS_Product_Mar-{25-i}": float(eps[f"Mar{25-i}"]) * float(weights[f"Mar{25-i}"]),
                    f"Issue_Size_(Rs_Cr)_-_Public_Mar-{25-i}": fresh_issue_size_cr,
                    f"ofs_issue_size_cr{i+1}": ofs_issue_size_cr,
                    f"total_issue_size_cr{i+1}": total_issue_size_cr,
                    f"Post_Issue_Valuation_of_Company_(Rs_Cr)_Mar-{25-i}": post_issue_valuation_cr,
                    f"PE_Ratio_(2025_EPS)_Mar-{25-i}": pe_ratio,
                    f"PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-{25-i}": pe_post_ratio,
                    f"PE_Ratio_Weighted_Average_EPS_Mar-{25-i}": pe_weighted,
                    "bv_pre": bv_pre,
                    f"Pre_Issue_PBV_Mar-{25-i}": pbv_pre,
                    f"Post_Issue_Networth_(`_cr)_Mar-{25-i}": post_issue_networth_lakh,
                    f"Post_Issue_Book_Value_Mar-{25-i}": bv_post,
                    f"Post_Issue_PBV_Mar-{25-i}": pbv_post,
                    "Subtotal_(A)_Shareholding_Pattern_Post_IPO_Number_of_Shares" : int(promoter_post_issue_shares),
                    # "Promoter_group_Shareholding_Pattern_Post_IPO_Number_of_Shares" : promoter_post_issue_shares,
                    "Subtotal_(A)_Shareholding_Pattern_Post_IPO_Percentage" : promoter_post_issue_pct,
                    # "Promoter_group_Shareholding_Pattern_Post_IPO_Percentage" : promoter_post_issue_pct,
                    "Total_(A+B)_Shareholding_Pattern_Post_IPO_Number_of_Shares" : int(post_issue_shares),
                    "Subtotal_(B)_Shareholding_Pattern_Post_IPO_Number_of_Shares" : int(public_post_issue_shares),
                    "Subtotal_(B)_Shareholding_Pattern_Post_IPO_Percentage" : public_post_issue_pct,
                    "AIF/Individuals/corporates_Shareholding_Pattern_Post_IPO_Number_of_Shares" : int(aif_post_issue_shares),
                    "AIF/Individuals/corporates_Shareholding_Pattern_Post_IPO_Percentage" : aif_post_issue_pct,
                    "IPO_shareholders_Shareholding_Pattern_Post_IPO_Percentage" : ipo_shareholders_pct,
                    "Total_(A+B)_Shareholding_Pattern_Post_IPO_Percentage" : "100%",
                    "Total_(A+B)_Shareholding_Pattern_Pre-IPO_Number_of_Shares" : int(total_pre_issue_shares),
                    "Total_(A+B)_Shareholding_Pattern_Pre-IPO_Percentage" : "100%"
                },
            })
    else:
        results.append({
            "placeholders": {
                "Share_Holding_Pre_Issue": pre_issue_shares,
                "issue_size": fresh_issue_shares,
                "ofs_promoter_shares": ofs_promoter_shares,
                "Subtotal_(A)_Shareholding_Pattern_Pre-IPO_Number_of_Shares": int(promoter_pre_issue_shares),
                "weighted_avg_eps" : wtd_avg_eps,
                # f"Price_(Rs)_Mar-{25-i}": issue_price,
                "pre_issue_networth_lakh": pre_issue_networth_lakh,
                "issue_expenses_lakh": issue_expenses_lakh,
                "pat_lakh": pat_lakh,
                "1._PAT" : pat_latest,
                "1._Net_worth" : net_worth_latest,
                "1._P/E" : pe_latest,
                "1._PBV" : pbv_self,
                "1._EPS" : eps_latest,
                "1._Issue_Price" : issue_price_avg,
                "PAT_(`_in_lakh)_Mar-25" : pat_lakh_25,
                "PAT_(`_in_lakh)_Mar-24" : pat_lakh_24,
                "PAT_(`_in_lakh)_Mar-23" : pat_lakh_23,
                "Peer_1._Net_worth": net_worth_1,
                "Peer_1._Issue_Price": issue_price_p1,
                "Peer_1._PBV": pbv_p1,
                "Peer_2._Net_worth": net_worth_2,
                "Peer_2._Issue_Price": issue_price_p2,
                "Peer_2._PBV": pbv_p2,
                # f"Weight_Mar-{25-i}": weights[f"Mar{25-i}"],
                # f"EPS_Mar-{25-i}" : eps[f"Mar{25-i}"],
                # f"EPS_after_adjusting_Bonus_Mar-{25-i}" : eps[f"Mar{25-i}"],
                # f"EPS_after_Proposed_Issue_Mar-{25-i}" : eps_after_issue[f"Mar{25-i}"],
                # f"PE_Ratio_Weighted_Average_EPS_Mar-{25-i}" : wtd_avg_eps,
                # f"EPS_Product_Mar-{25-i}": float(eps[f"Mar{25-i}"]) * float(weights[f"Mar{25-i}"]),
                # f"Issue_Size_(Rs_Cr)_-_Public_Mar-{25-i}": fresh_issue_size_cr,
                # f"ofs_issue_size_cr{i+1}": ofs_issue_size_cr,
                # f"total_issue_size_cr{i+1}": total_issue_size_cr,
                # f"Post_Issue_Valuation_of_Company_(Rs_Cr)_Mar-{25-i}": post_issue_valuation_cr,
                # f"PE_Ratio_(2025_EPS)_Mar-{25-i}": pe_latest,
                # f"PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-{25-i}": pe_post_issue_latest,
                # f"PE_Ratio_Weighted_Average_EPS_Mar-{25-i}": pe_weighted,
                # "bv_pre": bv_pre,
                # f"Pre_Issue_PBV_Mar-{25-i}": pbv_pre,
                # f"Post_Issue_Networth_(`_cr)_Mar-{25-i}": post_issue_networth_lakh,
                # f"Post_Issue_Book_Value_Mar-{25-i}": bv_post,
                # f"Post_Issue_PBV_Mar-{25-i}": pbv_post,
                "Subtotal_(A)_Shareholding_Pattern_Post_IPO_Number_of_Shares" : int(promoter_post_issue_shares),
                # "Promoter_group_Shareholding_Pattern_Post_IPO_Number_of_Shares" : promoter_post_issue_shares,
                "Subtotal_(A)_Shareholding_Pattern_Post_IPO_Percentage" : promoter_post_issue_pct,
                # "Promoter_group_Shareholding_Pattern_Post_IPO_Percentage" : promoter_post_issue_pct,
                "Total_(A+B)_Shareholding_Pattern_Post_IPO_Number_of_Shares" : int(post_issue_shares),
                "Subtotal_(B)_Shareholding_Pattern_Post_IPO_Number_of_Shares" : int(public_post_issue_shares),
                "Subtotal_(B)_Shareholding_Pattern_Post_IPO_Percentage" : public_post_issue_pct,
                "AIF/Individuals/corporates_Shareholding_Pattern_Post_IPO_Number_of_Shares" : int(aif_post_issue_shares),
                "AIF/Individuals/corporates_Shareholding_Pattern_Post_IPO_Percentage" : aif_post_issue_pct,
                "IPO_shareholders_Shareholding_Pattern_Post_IPO_Percentage" : ipo_shareholders_pct,
                "Total_(A+B)_Shareholding_Pattern_Post_IPO_Percentage" : "100%",
                "Total_(A+B)_Shareholding_Pattern_Pre_IPO_Number_of_Shares" : int(total_pre_issue_shares),
                "Total_(A+B)_Shareholding_Pattern_Pre_IPO_Percentage" : "100%"
            },
        })
        
    return results # type: ignore

def fill_html2(html_content: str, json_content: str) -> str:
    print("*"*100)
    print(json_content)
    print("*"*100)
    try:
        data = json.loads(json_content).get("extracted_json", [])
        table_data_list = [list(d["placeholders"].values()) for d in data if "placeholders" in d]
        text_data = []
        for d in data:
            if "placeholders" not in d:
                for _, v in d.items():
                    text_data.append(v)

        soup = BeautifulSoup(html_content, "html.parser")
        for tag in soup.find_all(["p", "li", "span", "strong", "td", "th"]):
            tag.string = tag.get_text()
        soup.smooth()    
        for tag in soup.find_all(["p", "li", "span", "strong", "td", "th"]):
            if tag.name == "p":
                tag["style"] = "font-family:Rupee Foradian;font-size:11pt"
            else:
                tag["style"] = "font-family:Rupee Foradian;font-size:10pt"

        tables = soup.find_all("table")
        for t_idx, table in enumerate(tables):
            if t_idx >= len(table_data_list):
                break
            values = table_data_list[t_idx]
            vi = 0
            for tr in table.find_all("tr"):
                cells = tr.find_all(["td", "th"])
                if not cells:
                    continue
                texts = [c.get_text(strip=True) for c in cells]
                if len(texts) == 1:continue
                start_idx = 1 if len(texts) >= 2 and texts[0] == "" and texts[1] != "" else 0
                for i in range(start_idx, len(cells)):
                    if vi < len(values) and texts[i] == "":
                        cells[i].string = to_crores(values[vi])
                        vi += 1
        
        text_idx = 0

        for i in range(len(text_data)):
            if text_data[i] == "" or text_data[i].lower() == "na":
                text_data[i] = "Couldn't be Captured"
        
        for node in soup.descendants:
            if not isinstance(node, NavigableString):
                continue
            
            parent = node.parent
            if parent and parent.name == "table":
                continue
            
            text = str(node)
            if "[" not in text or "]" not in text:
                continue
            
            parts = re.split(r'(\[[^\[\]]*\])', text)
            new_text = ""
            
            for part in parts:
                if part.startswith("[") and part.endswith("]"):
                    if text_idx< len(text_data):
                        new_text += text_data[text_idx]
                        text_idx += 1
                    else:
                        new_text += "Couldn't be Captured."
                else:
                    new_text += part

            if new_text is None or new_text == "" or new_text.lower() =="na":
                new_text = "Couldn't be Captured."
        
            node.replace_with(NavigableString(new_text))
            
        return str(soup)


    except Exception as e:
        return f"<html><body><h3>Error processing HTML: {e}</h3></body></html>"

# input_json = [
#     {
#         "placeholders": {
#             "Name_of_Issuer_Company": "Shyam Dhani Industries Limited",
#             "MSME_status/_Udyam_Status": "Couldn't be captured",
#             "Date_of_Incorporation": "October 19,2010",
#             "Commencement_of_Operations": "Couldn't be captured",
#             "Address_of_Registered_Office": "F-438A, Road No. 12, VKIA, Jaipur, Rajasthan, 302013, India",
#             "Address_of_Operating_facilities": "Khasra No. 06/1067, Manpura Road, Jatawali, near Delhi bypass, Tehsil – Chomu, Jaipur, Rajasthan"
#         }
#     },
#     {
#         "Brief_Background_of_company_and_promoters": "Shyam Dhani Industries Limited is engaged in manufacturing and processing of 164 types/varieties of spices under the brand name 'SHYAM'. The company also trades and distributes Grocery Products and a diverse range of Herbs and seasonings. Originally incorporated as a Private Limited Company on October 19, 2010, it was converted to Public Limited Company on August 20, 2024. The promoters are Mr. Ramawtar Agarwal (Chairman and Managing Director with PhD in Food Industry from Maryland State University), Mrs. Mamta Devi Agarwal (Whole Time Director), and Mr. Vithal Agarwal (Whole Time Director)."
#     },
#     {
#         "Details_about_DRHP_filing_date_approval_date": "The Draft Red Herring Prospectus was filed on August 30, 2025, and approved by the Board on August 30, 2025. The Red Herring Prospectus was approved by the Board on December 09, 2025. The company received In-Principal approval from NSE on November 24, 2025."
#     },
#     {
#         "Details_about_Exchange_on_which_company_is_going_to_be_listed_Merchant_banker_and_Registrar.": "The company will be listed on the EMERGE Platform of National Stock Exchange of India Limited (NSE EMERGE). Book Running Lead Manager: Holani Consultants Private Limited. Registrar to the Issue: Bigshare Services Private Limited. Market Maker: Holani Consultants Private Limited."
#     },
#     {
#         "placeholders": {
#             "IPO_Date": "Monday, December 22, 2025 to Wednesday, December 24, 2025",
#             "Listing_Date": "Couldn't be Sourced",
#             "Face_Value": "₹ 10",
#             "Issue_Price_Band": "10,12",
#             "Lot_Size": "Couldn't be Sourced",
#             "Sale_Type": "Fresh Issue",
#             "Issue_size": "Up to 54,98,000 Equity Shares",
#             "Issue_Type": "Initial Public Offering",
#             "Listing_At": "EMERGE Platform of NSE",
#             "Share_Holding_Pre_Issue": "1,51,58,000 Equity Shares",
#             "Share_Holding_Post_Issue": "Couldn't be Sourced"
#         }
#     },
#     {
#         "placeholders": {
#             "QIB_Shares_Offered_Shares_Offered": "26,04,000 (47.36%)",
#             "Retail_Shares_Offered_Shares_Offered": "18,28,000 (33.25%)",
#             "NII_Shares_Offered_Shares_Offered": "7,86,000 (14.30%)"
#         }
#     },
#     {
#         "placeholders": {
#             "Assets_31_March_2025": "82.47",
#             "Assets_31_Mar_2024": "52.84",
#             "Assets_31_Mar_2023": "27.51",
#             "Total_Income_31_March_2025": "124.75",
#             "Total_Income_31_Mar_2024": "107.64",
#             "Total_Income_31_Mar_2023": "68.10",
#             "Profit_After_Tax_31_March_2025": "8.04",
#             "Profit_After_Tax_31_Mar_2024": "6.30",
#             "Profit_After_Tax_31_Mar_2023": "2.92",
#             "Net_Worth_31_March_2025": "23.61",
#             "Net_Worth_31_Mar_2024": "1556 lakhs",
#             "Net_Worth_31_Mar_2023": "9.26",
#             "Reserves_and_Surplus_31_March_2025": "8.73",
#             "Reserves_and_Surplus_31_Mar_2024": "14.42",
#             "Reserves_and_Surplus_31_Mar_2023": "8.12",
#             "Total_Borrowing_31_March_2025": "47.24",
#             "Total_Borrowing_31_Mar_2024": "24.45",
#             "Total_Borrowing_31_Mar_2023": "12.75"
#         }
#     },
#     {
#         "Brief_about_company's_projections_profitability_etc": "Shyam Dhani Industries Ltd.'s revenue increased by 16% and profit after tax (PAT) rose by 28% between the financial year ending with March 31, 2025 and March 31, 2024. The company marked steady growth in its top and bottom lines for the reported periods."
#     },
#     {
#         "placeholders": {
#             "Promoter_Shareholding_Pattern_Pre-IPO_Nos": "1,41,40,750",
#             "Promoter_Shareholding_Pattern_Pre-IPO_%": "93.30%",
#             "Promoter_Shareholding_Pattern_Post_IPO_Nos": "Couldn't be Sourced",
#             "Promoter_Shareholding_Pattern_Post_IPO_%": "Couldn't be Sourced",
#             "Promoter_group_Shareholding_Pattern_Pre-IPO_Nos": "7,31,250",
#             "Promoter_group_Shareholding_Pattern_Pre-IPO_%": "4.81",
#             "Promoter_group_Shareholding_Pattern_Post_IPO_Nos": "Couldn't be Sourced",
#             "Promoter_group_Shareholding_Pattern_Post_IPO_%": "Couldn't be Sourced",
#             "Subtotal_(A)_Shareholding_Pattern_Pre-IPO_Nos": "1,48,72,000",
#             "Subtotal_(A)_Shareholding_Pattern_Pre-IPO_%": "98.11%",
#             "Subtotal_(A)_Shareholding_Pattern_Post_IPO_Nos": "Couldn't be Sourced",
#             "Subtotal_(A)_Shareholding_Pattern_Post_IPO_%": "Couldn't be Sourced",
#             "AIF/Individuals/corporates_Shareholding_Pattern_Pre-IPO_Nos": "2,86,000",
#             "AIF/Individuals/corporates_Shareholding_Pattern_Pre-IPO_%": "1.89%",
#             "AIF/Individuals/corporates_Shareholding_Pattern_Post_IPO_Nos": "Couldn't be Sourced",
#             "AIF/Individuals/corporates_Shareholding_Pattern_Post_IPO_%": "Couldn't be Sourced",
#             "IPO_shareholders_Shareholding_Pattern_Pre-IPO_Nos": "0",
#             "IPO_shareholders_Shareholding_Pattern_Pre-IPO_%": "0%",
#             "IPO_shareholders_Shareholding_Pattern_Post_IPO_Nos": "54,98,000",
#             "IPO_shareholders_Shareholding_Pattern_Post_IPO_%": "Couldn't be Sourced",
#             "Subtotal_(B)_Shareholding_Pattern_Pre-IPO_Nos": "2,86,000",
#             "Subtotal_(B)_Shareholding_Pattern_Pre-IPO_%": "1.89%",
#             "Subtotal_(B)_Shareholding_Pattern_Post_IPO_Nos": "Couldn't be Sourced",
#             "Subtotal_(B)_Shareholding_Pattern_Post_IPO_%": "Couldn't be Sourced",
#             "Total_(A+B)_Shareholding_Pattern_Pre-IPO_Nos": "1,44,26,750",
#             "Total_(A+B)_Shareholding_Pattern_Pre-IPO_%": "95.19%",
#             "Total_(A+B)_Shareholding_Pattern_Post_IPO_Nos": "Couldn't be Sourced",
#             "Total_(A+B)_Shareholding_Pattern_Post_IPO_%": "Couldn't be Sourced"
#         }
#     },
#     {
#         "placeholders": {
#             "DRHP_should_have_been_approved_by_any_Stock_Exchange_/_SEBI._Compliance": "",
#             "Should_have_availed_at_least_one_Debt_products_from_Banks/FIs_and_servicing_it_for_at_least_3_years_prior_to_the_date_of_filing_of_DRHP._and_/_or_should_have_availed_funding_from_any_SEBI_registered_AIFs_at_least_1_year_prior_to_the_date_of_filing_of_DRHP._Compliance": "",
#             "Net_worth_>_₹_20_crore_Compliance": "",
#             "EBIDTA_margin_-_₹2.00_cr._or_>_5%_(whichever_is_more)_for_at_least_two_out_of_the_three_most_recent_financial_years_Compliance": "",
#             "Total_Income_>_₹50_crore_Compliance": "",
#             "Net_Tangible_Assets_>_₹_2.50_crore_Compliance": "",
#             "Profitability_(total_PBT_of_3_years)_>_₹_10_crore_Compliance": "",
#             "Exposure_cap_per_MSME:_Minimum:_₹_1_crore_Maximum:_Up_to_₹_20_crore._Note:_Anchor_investment_shall_not_exceed_50%_of_the_Anchor_portion_of_the_Issue_or_10%_of_the_post_issue_paid-up_capital_of_the_issuer_company_whichever_is_lower._Compliance": "",
#             "Minimum_Promoters_Shareholding_(Post_Issue)-_60%_In_case_where_the_Company_has_received_any_investment_from_a_SEBI_registered_AIF_the_Investment_Committee_may_take_a_suitable_view_on_the_minimum_level_of_Promoter_Shareholding_below_60%_since_the_AIF_would_want_to_partially_/_fully_exit_in_the_proposed_IPO_/_Offer_for_Sale_(OFS)._Compliance": "",
#             "Complied_with_Minimum_₹_2_Cr._Bid_criteria_of_SEBI_Compliance": "",
#             "Complied_with_RBI's_guideline_limiting_post-issue_shareholding_exposure_to_a_maximum_of_10%._Compliance": ""
#         }
#     },
#     {
#         "Brief_about_industry": "Shyam Dhani Industries Limited is an ISO-certified company and manufacturer, exporter, wholesaler, and supplier of Premium Spices, Spices Powder, Whole Spices, etc. They trade and distribute Grocery Products like Black Salt, Rock Salt, Rice, Poha, Kasuri Methi, and a variety of Herbs and seasonings. The company processes 163 varieties of spices, including Ground, Blend, and Whole Spices under the brand name 'SHYAM'."
#     },
#     {
#         "placeholders": {
#             "1_Items": "Funding incremental working capital requirements",
#             "1_Amount_(₹_crore)": "13.26",
#             "1_%": "Couldn't be captured",
#             "2_Items": "Repayment/Pre-payment of outstanding borrowings",
#             "2_Amount_(₹_crore)": "10.00",
#             "2_%": "Couldn't be captured",
#             "3_Items": "Brand Creation and Marketing Expenses",
#             "3_Amount_(₹_crore)": "6.36",
#             "3_%": "Couldn't be captured",
#             "4_Items": "Capital Expenditure for machinery purchase",
#             "4_Amount_(₹_crore)": "1.63",
#             "4_%": "Couldn't be captured",
#             "Total_Items": "Total Objects",
#             "Total_Amount_(₹_crore)": "Couldn't be captured",
#             "Total_%": "100%"
#         }
#     },
#     {
#         "Brief_about_basis_of_valuation": "The company's valuation is based on book building process. Industry peer group P/E ratio ranges from 3.06 to 13.56 with industry composite of 8.31. The company has shown consistent profitability with improving margins and strong financial performance."
#     },
#     {
#         "placeholders": {
#             "PAT_(`_in_lakh)_Mar-25": "804.16",
#             "PAT_(`_in_lakh)_Mar-24": "630.29",
#             "PAT_(`_in_lakh)_Mar-23": "292.40",
#             "EPS_Mar-25": "5.31",
#             "EPS_Mar-24": "",
#             "EPS_Mar-23": "",
#             "EPS_after_adjusting_Bonus_Mar-25": "",
#             "EPS_after_adjusting_Bonus_Mar-24": "",
#             "EPS_after_adjusting_Bonus_Mar-23": "",
#             "EPS_after_Proposed_Issue_Mar-25": "4.07",
#             "EPS_after_Proposed_Issue_Mar-24": "",
#             "EPS_after_Proposed_Issue_Mar-23": "",
#             "Weight_Mar-25": "",
#             "Weight_Mar-24": "",
#             "Weight_Mar-23": "",
#             "EPS_Product_Mar-25": "",
#             "EPS_Product_Mar-24": "",
#             "EPS_Product_Mar-23": "",
#             "weighted_avg_eps": "",
#             "PE_Ratio_(2025_EPS)_Mar-25": "13.19",
#             "PE_Ratio_(2025_EPS)_Mar-24": "",
#             "PE_Ratio_(2025_EPS)_Mar-23": "",
#             "PE_Ratio_Weighted_Average_EPS_Mar-25": "",
#             "PE_Ratio_Weighted_Average_EPS_Mar-24": "",
#             "PE_Ratio_Weighted_Average_EPS_Mar-23": "",
#             "PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-25": "17.21",
#             "PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-24": "",
#             "PE_Ratio_Post_issue_based_on_FY_2025_Earning_Mar-23": "",
#             "Price_(Rs)_Mar-25": "",
#             "Price_(Rs)_Mar-24": "",
#             "Price_(Rs)_Mar-23": "",
#             "Issue_Size_(Rs_Cr)_-_Public_Mar-25": "",
#             "Issue_Size_(Rs_Cr)_-_Public_Mar-24": "",
#             "Issue_Size_(Rs_Cr)_-_Public_Mar-23": "",
#             "Post_Issue_Valuation_of_Company_(Rs_Cr)_Mar-25": "",
#             "Post_Issue_Valuation_of_Company_(Rs_Cr)_Mar-24": "",
#             "Post_Issue_Valuation_of_Company_(Rs_Cr)_Mar-23": "",
#             "Pre_Issue_PBV_Mar-25": "",
#             "Pre_Issue_PBV_Mar-24": "",
#             "Pre_Issue_PBV_Mar-23": "",
#             "Post_Issue_Book_Value_Mar-25": "",
#             "Post_Issue_Book_Value_Mar-24": "",
#             "Post_Issue_Book_Value_Mar-23": "",
#             "Post_Issue_PBV_Mar-25": "",
#             "Post_Issue_PBV_Mar-24": "",
#             "Post_Issue_PBV_Mar-23": "",
#             "Post_Issue_Networth_(`_cr)_Mar-25": "",
#             "Post_Issue_Networth_(`_cr)_Mar-24": "",
#             "Post_Issue_Networth_(`_cr)_Mar-23": ""
#         }
#     },
#     {
#         "placeholders": {
#             "1._Company_Name": "",
#             "1._FV": "",
#             "1._Net_worth": "",
#             "1._Sales": "",
#             "1._PAT": "",
#             "1._EPS": "",
#             "1._Issue_Price": "",
#             "1._P/E": "",
#             "1._B.V_Pre-Issue": "",
#             "1._PBV": "",
#             "1._RONW(%)": "",
#             "Peer_1._Company_Name": "",
#             "Peer_1._FV": "",
#             "Peer_1._Net_worth": "",
#             "Peer_1._Sales": "",
#             "Peer_1._PAT": "",
#             "Peer_1._EPS": "",
#             "Peer_1._Issue_Price": "",
#             "Peer_1._P/E": "",
#             "Peer_1._B.V_Pre-Issue": "",
#             "Peer_1._PBV": "",
#             "Peer_1._RONW(%)": "",
#             "Peer_2._Company_Name": "",
#             "Peer_2._FV": "",
#             "Peer_2._Net_worth": "",
#             "Peer_2._Sales": "",
#             "Peer_2._PAT": "",
#             "Peer_2._EPS": "",
#             "Peer_2._Issue_Price": "",
#             "Peer_2._P/E": "",
#             "Peer_2._B.V_Pre-Issue": "",
#             "Peer_2._PBV": "",
#             "Peer_2._RONW(%)": ""
#         }
#     },
#     {
#         "total_issue_size_price": "Up to 54,98,000 Equity Shares of face value ₹ 10 each"
#     },
#     {
#         "Number_of_shares": "54,98,000 Equity Shares"
#     },
#     {
#         "Number_of_lots": "Couldn't be captured"
#     },
#     {
#         "Number_of_shares": "54,98,000 Equity Shares"
#     },
#     {
#         "about_promoters_experience": "Mr. Ramawtar Agarwal has PhD in Food Industry from Maryland State University and has been proprietor of Shyam Dhani Industries since March 1995. Mrs. Mamta Devi Agarwal has experience in trading business. Mr. Vithal Agarwal joined the company in December 2017. All promoters have adequate experience in the line of business."
#     },
#     {
#         "about_consistent_in_profitability": "The company has shown consistent profitability with PAT growing from ₹ 292.40 lakhs in FY23 to ₹ 804.16 lakhs in FY25. EBITDA margins have improved from 10.12% to 13.58% and PAT margins from 5.86% to 6.59% over the same period."
#     },
#     {
#         "about_entry_barriers": "The spice industry has moderate entry barriers including need for quality certifications like FSSAI, AGMARK, ISO 22000:2018, brand building requirements, distribution network establishment, and compliance with food safety regulations."
#     },
#     {
#         "any_other_rationale": "The company has diversified product portfolio of 164 varieties of spices, strong presence in modern trade, export capabilities with FDA registration for USA, and strategic brand ambassador appointment of Ms. Preity Zinta for market expansion."
#     },
#     {
#         "About_Lead_and_Co-lead_merchant_bankers_their_past_track_record": "Book Running Lead Manager is Holani Consultants Private Limited. The document mentions other regulatory and statutory disclosures and price information of past issues handled, but specific track record details and percentage of issues trading positive/negative on 180th day are not clearly captured in the provided text."
#     },
#     {
#         "Brief_about_Market_Maker_and_their_track_record": "Market Maker is Holani Consultants Private Limited with SEBI Registration No. INZ000299835. They will provide 2-way quotes for 75% of the time with spread not more than 10% and minimum depth of quote ₹ 1,00,000 for three years from listing date."
#     },
#     {
#         "On_the_basis_of_Credibility_of_issuer_Mechant_banker_potential_of_returns_Good_investors_backing_etc": "The company has strong financial performance with consistent growth, experienced promoters, diversified product portfolio, and established market presence. The issue is 100% underwritten and has received in-principle approval from NSE."
#     },
#     {
#         "Date": "December 22, 2025 to December 24, 2025"
#     },
#     {
#         "Offer_for_sale_size": "Not applicable - entire issue constitutes fresh issue"
#     },
#     {
#         "placeholders": {
#             "Compliance_of_KYC_guidelines": "Couldn't be captured",
#             "KYC_Risk_Score": "Couldn't be captured",
#             "Verification_of_defaulters'_lists_Watchout_Investor_website_Income_Tax_defaulter_Search_etc.": "Company and promoters are not declared as wilful defaulters or fraudulent borrowers",
#             "Verification_of_CIBIL_reports_For_Company_and_Promoters": "Couldn't be captured",
#             "JOCATA_Score": "Couldn't be captured",
#             "Market_Inquiries/_Feedback_from_references": "Couldn't be captured",
#             "Connected_Lending": "Couldn't be captured",
#             "Pending_court_cases_initiated_by_other_banks/_FIs_against_the_company_promoters/director_etc_if_any": "Outstanding material civil litigation including trademark infringement case",
#             "Pending_court_cases_initiated_by_third_parties_(Debtors/_creditors/_competitors)_against_the_company_promoters/director_etc_if_any": "Everest Food Products case regarding trademark infringement",
#             "Certification/_Compliances": "Company has various certifications including FSSAI, AGMARK, ISO 22000: 2018, APEDA",
#             "If_any_Adverse_Audit_Remarks_in_last_3_years_audited_Annual_Reports": "No auditor qualifications in Restated Financial Statements",
#             "Position_of_statutory_dues": "Delays in filing certain forms and payment of employee-related statutory dues in the past",
#             "Contingent_Liability_if_any": "₹ 8.96 Lakhs as of September 30,2025",
#             "BO_Remarks_(in_case_of_SIDBI_customer)": "Couldn't be captured",
#             "Remarks_of_MSME_Equity_Cell:": "Couldn't be captured",
#             "(a)_Credibility_of_Issuer": "Strong financial performance with consistent growth and experienced management",
#             "(b)_Credibility_of_Merchant_Banker": "Holani Consultants Private Limited is SEBI registered Category I Merchant Banker",
#             "(c)_Liability_on_SIDBI": "Couldn't be captured",
#             "(d)_Proposed_Investment": "Fresh issue of up to 54,98,000 equity shares"
#         }
#     }
# ]

def fill_html_with_json(html_content, data):
    print("initital_llm_data ==>", data)
    print("\n\n")
    try:
        print("Data from llm ===>",data)
        print("\n\n----------\n\n")
        calc_list1 = {}
        for d in data:
            if "placeholders" in d:
                for k,v in d["placeholders"].items():
                    calc_list1[k] = v
        calc_list2 = {}
        for d in data:
            if ("placeholders" not in d) and ("calculations" not in d):
                for i,j in d.items():
                    calc_list2[f'{i}'] = j                
        calc_list = calc_list1 | calc_list2
        s = calc_list["Issue_Price_Band"]
        numbers = re.findall(r"\d+(?:\.\d+)?",s)
        ax = True
        if len(numbers) == 2:
            calc_list["issue_price1"] = numbers[0]
            calc_list["issue_price3"] = numbers[1]
        elif len(numbers) == 1:
            calc_list["issue_price1"] = numbers[0]
            calc_list["issue_price3"] = numbers[0]
        else :
            calc_list["issue_price1"] = 0
            calc_list["issue_price3"] = 0
            ax=False

        for k,v in calc_list.items():
            if isinstance(v,(list,dict)) and k != "weights":
                calc_list[k] = 0   #add a list, dict checker just in case.
        
        # print(f"Inside html_to_json and prinitng calc_list====>{calc_list}")
        
        calc_results = []
        payload = SimulationRequest(**calc_list)
        result = compute_issue_simulation(ax,payload)
        calc_results.append(result)
        calc = []
        for a in calc_results:
            calc = a
        changer = {}
        for d in calc:
            for k,v in d["placeholders"].items():
                changer[k] = v
        for k,v in changer.items():
            for d in data:
                if "placeholders" in d:
                    if k in d["placeholders"].keys():
                        if type(v) == int or type(v) == float:
                            d["placeholders"][k] = round(v,2)
                        if type(v) == str:
                            d["placeholders"][k] = v
                        
        print("Changer===>",changer) 
        print("\n--------------\n")
        print("CHANGED DATA====>",data)
        print("\n--------------\n")                                                               
        table_data_list=[list(d["placeholders"].values()) for d in data if "placeholders" in d]
        text_data=[]
        for d in data:
            if ("placeholders" not in d) and ("calculations" not in d):
                for _,v in d.items(): 
                    text_data.append(v)
        soup=BeautifulSoup(html_content,"html.parser")
        for tag in soup.find_all(["p","li","span","strong","td","th"]):
            tag.string = tag.get_text()
        soup.smooth()
        for tag in soup.find_all(["p","li","span","strong","td","th"]):
            if tag.name=="p": tag["style"]="font-family:Rupee Foradian;font-size:11pt"
            else: tag["style"]="font-family:Rupee Foradian;font-size:10pt"
        tables=soup.find_all("table")
        for t_idx,table in enumerate(tables):
            if t_idx>=len(table_data_list): break
            values=table_data_list[t_idx]
            vi=0
            for tr in table.find_all("tr"):
                cells=tr.find_all(["td","th"])
                if not cells: 
                    continue
                texts=[c.get_text(strip=True) for c in cells]
                if len(texts) == 1: continue
                if len(texts)>=2 and texts[0]=="" and texts[1]=="": 
                    continue
                start_idx=0
                if len(texts)>=2 and texts[0]=="" and texts[1]!="": 
                    start_idx=1
                for i in range(start_idx,len(cells)):
                    if vi>= len(values):
                        break
                    cell = cells[i]
                    cell_text = cell.get_text(strip=True)
                    
                    if cell_text != "":
                        continue
                    
                    cell.string = to_crores(values[vi])
                    vi+=1
        
        text_idx = 0
        text_chk = 0
        # print("=========="*14)
        # print("Printing text data ",text_data)
        # print("=========="*14)
        for i in range(len(text_data)):
            if text_data[i] == "" or text_data[i].lower() == "na" or text_data[i]==" ":
                text_data[i] = "Couldn't be Captured"
        
        # print(text_data)
        for node in soup.descendants:
            if not isinstance(node, NavigableString):
                continue
            
            parent = node.parent
            if parent and parent.name == "table":
                continue
            
            text = str(node)
            if "[" not in text or "]" not in text:
                continue
            
            if "[●]" in text:
                continue
            
            parts = re.split(r'(\[[^\[\]]*\])', text)
            new_text = ""
            
            for part in parts:
                if part.startswith("[") and part.endswith("]"):
                    if text_idx< len(text_data):
                        new_text += text_data[text_idx]
                        text_idx += 1
                    else:
                        new_text += "Couldn't be Captured."
                    text_chk+=1
                else:
                    new_text += part
                    
            if new_text is None or new_text == "" or new_text.lower() == "na":
                new_text = "Couldn't be Captured."
                    
            node.replace_with(NavigableString(new_text))
        
        hex = "#DA4108"
            
        pattern = re.compile(r"Couldn't be Captured", re.IGNORECASE)
        
        for text_node in soup.find_all(string=pattern):
            if text_node.parent.name in ["script","style"]:
                continue
            
            new_html = pattern.sub(
                lambda m : f'<span style="color:{hex};">{m.group()}</span>',
                text_node
            )

            text_node.replace_with(BeautifulSoup(new_html,"html.parser"))
        
        ptrn = re.compile(r"Business to fill", re.IGNORECASE)
        
        for text_node in soup.find_all(string=ptrn):
            if text_node.parent.name in ["script","style"]:
                continue
            
            text = text_node.strip()
            
            new_html2 = (
                f'<span style="background-color: yellow; color: black;">'
                f'{text}'
                f'</span>'
            )
            
            text_node.replace_with(BeautifulSoup(new_html2,"html.parser"))

        return str(soup)
    
    except Exception as e:
        return f"<html><body><h3>Error processing with the the HTML: {e}</h3></body></html>"


# driver_service_location = "edge.exe"

# def create_driver():
#     options = Options()
#     options.add_argument("--headless=new")
#     options.add_argument("--no-sandbox")
#     options.add_argument("--disable-dev-shm-usage")
#     options.add_argument("--disable-gpu")
#     options.add_argument("--disable-blink-features=AutomationControlled")
#     options.page_load_strategy = "eager"
#     options.add_argument("--ignore-certificate-errors")
#     options.add_argument(f'--proxy-server=http://{PROXY}')

#     service = Service("/usr/local/bin/msedgedriver")

#     driver = webdriver.Edge(service=service, options=options)
    
    
#     return driver

# -------------------------
# SAFE PAGE LOAD
# -------------------------
# def safe_get(driver,url,wait=None):
#     for atmp in range(1,max_retry+1):
#         try:
#             driver.set_page_load_timeout(page_load_timeout)
#             driver.get(url)

#             if wait:
#                 WebDriverWait(driver,page_load_timeout).until(
#                     EC.presence_of_element_located(wait)
#                 )
#             return True

#         except TimeoutException:
#             if atmp == max_retry:
#                 raise
#             driver.execute_script("window.stop();")
#     return False            


# def get_ipo_url(query):
#     driver = None
    
#     try:
#         driver = create_driver()
#         url = "https://www.chittorgarh.com/report/ipo-in-india-list-main-board-sme/82/mainboard/"
#         safe_get(driver,url,wait=(By.CSS_SELECTOR,'button.search-btn-1'))
#         button = driver.find_element(By.CSS_SELECTOR,'button.search-btn-1')
#         driver.execute_script("arguments[0].click();",button)

#         search_input = WebDriverWait(driver,10).until(EC.visibility_of_element_located(
#         (By.CSS_SELECTOR,"input.form-control.rounded")))
#         search_input.clear()
#         start = time.time()
#         for ch in query:
#             if time.time() - start > 10:
#                 break
#             search_input.send_keys(ch)
#             time.sleep(0.2)
            
#             try:
#                 ul = driver.find_element(By.CSS_SELECTOR,"ul.list-group.position-absolute.shadow.rounded")

#                 result = ul.find_elements(By.TAG_NAME,"li")
#                 results = [li for li in result if li.is_displayed()]

#                 if len(results) == 1:
#                     driver.execute_script("arguments[0].click();",results[0])
#                     fnd = True
#                     break
#             except NoSuchElementException:
#                 pass            
        
#         WebDriverWait(driver,10).until(lambda d: d.current_url != url)
            
#         return driver.current_url

        
#     except Exception as e:
#         return f"Error: {e}"
    
#     finally:
#         if driver:
#             driver.quit()


# def get_data(url):
#     driver = None
#     try:
#         driver = create_driver()
#         driver.set_page_load_timeout(45)
#         # driver.get(url)
#         safe_get(driver,url)
#         time.sleep(3)
#         driver.execute_script("window.stop();")
        
#         # wait button safely
#         button = WebDriverWait(driver, 30).until(
#             EC.element_to_be_clickable(
#                 (By.XPATH, "//a[contains(@id,'anchorReadText')]")
#             )
#         )
#         driver.execute_script("arguments[0].click();", button)
        
#         jcontent = ""
#         accepted_div = []
#         titles = []

#         WebDriverWait(driver,30).until(
#             lambda d : d.execute_script("return document.readyState") == "complete"
#         )
        
#         divisions = driver.find_elements(By.XPATH,"//div[contains(@class,'card') and contains(@class,'p-3')]")
        
#         for div in divisions:
#             tbl = div.find_elements(By.TAG_NAME,'table')
#             lst = div.find_elements(By.TAG_NAME,'ul')
#             if tbl or lst:
#                 headings = div.find_elements(By.TAG_NAME,'h2')
#                 for heading in headings:
#                     if heading.text.strip():
#                         accepted_div.append(div)
#                         titles.append(heading.text.strip())
                    
#         for i in range(len(accepted_div)):
#             jcontent += f'---------{titles[i]}----------\n'
#             jcontent += accepted_div[i].text.strip()
#             jcontent += '\n\n'

#         return jcontent
    
#     except TimeoutException:
#         return f"Error: Page Load Timeout : {url}"
    
#     except Exception as e:
#         return f"Error extracting data: {e}"
    
#     finally:
#         if driver:
#             driver.quit()


# def full_exec(a):
#     try:
#         b = get_ipo_url(a)
#         if "Error extracting" in b:
#             return "url extraction error"
#         c = get_data(b)
#     except Exception as e:
#         c = str(e)
#     return c


import pymupdf as fitz
def extract_pdf_text(path):
    doc = fitz.open(path)
    blocks = []
    for page in doc:
        b = page.get_text("blocks")
        b.sort(key=lambda x: (x[1], x[0])) # type: ignore
        for blk in b:
            t = blk[4].strip()
            if t: blocks.append(t)
    return " ".join(blocks)

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning) # type: ignore

PROXIES = {
    "http": PROXY,
    "https": PROXY,
}


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Edg/121.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-IN,en;q=0.9",
    "Connection": "keep-alive",
}


# =========================================================
# SESSION CREATOR (retry + pooling + firewall safe)
# =========================================================
def create_session() -> requests.Session:
    session = requests.Session()

    retry = Retry(
        total=5,
        backoff_factor=1.5,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )

    adapter = HTTPAdapter(
        max_retries=retry,
        pool_connections=10,
        pool_maxsize=20
    )

    session.mount("http://", adapter)
    session.mount("https://", adapter)

    session.headers.update(HEADERS)
    session.proxies.update(PROXIES)


    return session


# =========================================================
# SAFE FETCH (robust error handling)
# =========================================================
def fetch_text(session: requests.Session, url: str, timeout: int = 30) -> str:
    try:
        resp = session.get(
            url,
            timeout=timeout,
            verify=False,
            allow_redirects=True
        )

        resp.raise_for_status()

        if not resp.text.strip():
            raise RuntimeError("Empty response received")

        return resp.text

    except requests.exceptions.Timeout:
        raise RuntimeError(f"Timeout while fetching: {url}")

    except requests.exceptions.ConnectionError:
        raise RuntimeError(f"Blocked by firewall/network: {url}")

    except requests.exceptions.SSLError:
        raise RuntimeError(f"SSL/TLS issue (proxy intercept?): {url}")

    except Exception as e:
        raise RuntimeError(f"{url} -> {e}")



# =========================================================
# PARSE SITEMAP
# =========================================================
# def parse_sitemap(xml_text: str):
#     root = ET.fromstring(xml_text)

#     if "}" in root.tag:
#         ns_uri = root.tag.split("}")[0].strip("{")
#         ns = {"sm": ns_uri}
#         loc_elems = root.findall(".//sm:loc", ns)
#     else:
#         loc_elems = root.findall(".//loc")

#     locs = [(e.text or "").strip() for e in loc_elems if (e.text or "").strip()]
#     kind = root.tag.split("}")[-1]

#     return kind, locs


# # =========================================================
# # FIND URL FROM SITEMAP
# # =========================================================
# def find_urls_by_keyword(keyword: str):
#     keyword = keyword.lower()
#     sitemap_url = "https://www.chittorgarh.com/sitemap.xml"

#     matches = []

#     with create_session() as session:

#         main_xml = fetch_text(session, sitemap_url)
#         kind, locs = parse_sitemap(main_xml)

#         if kind == "sitemapindex":
#             for sm_url in locs:
#                 time.sleep(0.2)
#                 try:
#                     sub_xml = fetch_text(session, sm_url)
#                     _, sub_locs = parse_sitemap(sub_xml)

#                     for u in sub_locs:
#                         if keyword in u.lower():
#                             matches.append(u)

#                 except Exception:
#                     continue

#         else:
#             for u in locs:
#                 if keyword in u.lower():
#                     matches.append(u)

#     # unique preserve order
#     return list(dict.fromkeys(matches))


# # =========================================================
# # GET IPO URL
# # =========================================================
# def get_url(query: str) -> str:
#     try:
#         q = query.lower().replace("limited", "").strip().replace(" ", "-")
#         keyword = f"ipo/{q}"

#         results = find_urls_by_keyword(keyword)

#         return results[0] if results else "Not found url"

#     except Exception as e:
#         return f"URL lookup error: {e}"
    
# # =========================================================
# # EXTRACT DATA FROM IPO PAGE (NO SELENIUM)
# # =========================================================
# def get_data_xml(url: str) -> str:
#     if url == "Not found url":
#         return "url not found."

#     try:
#         with create_session() as session:
#             html_text = fetch_text(session, url)

#         tree = html.fromstring(html_text)

#         divs = tree.xpath("//div[contains(@class,'card') and contains(@class,'p-3')]")

#         if not divs:
#             return "No data found. Page structure changed or blocked."

#         output = []

#         for div in divs:
#             heading = div.xpath(".//h2/text()")
#             text = div.text_content().strip()

#             if heading:
#                 output.append(f"-----{heading[0].strip()}-----\n{text}\n")

#         return "\n".join(output)

#     except Exception as e:
#         return f"Error extracting data: {e}"



# def full_exec_requests(query: str) -> str:
#     url = get_url(query)

#     if url.startswith("URL lookup error"):
#         return url

#     return get_data_xml(url)




@app.get("/get-backend-health")
def health():
    return {"status": "Backend services UP"}


# @app.post("/get-updated-url-data", response_class=PlainTextResponse)
# def url_get(query: str = Form(...)):
#     return full_exec_requests(query)

@app.post("/pdf-to-text2")
async def pdf_to_textx(file: UploadFile):
    if not file: return JSONResponse({"error":"no file"})
    with tempfile.NamedTemporaryFile(delete=False) as tmp:
        tmp.write(await file.read())
        path=tmp.name
    text=extract_pdf_text(path)
    os.remove(path)
    texts = text[:10000]
    return texts


# @app.post("/fill-html", response_class=HTMLResponse)
# async def fill_html(html: str = Form(...), json_data: str = Form(...)):
#     data=json.loads(json_data)["extracted_json"]
#     filled_html=fill_html_with_json_will_work(html,data)
#     return HTMLResponse(content=filled_html)


# This is the main endpoint being used right now
@app.post("/fill-html", response_class=HTMLResponse)
async def fill_html(payload:FillHtmlPayload):
    try:
        html = payload.template_html
        json_content = payload.data
        filled_html = fill_html_with_json(html, json_content)
        return HTMLResponse(content=filled_html)
    except Exception as e:
        print("Error in server: ",str(e))
        return HTMLResponse(content=f"Internal server error: {str(e)}")



@app.post("/fill-html2", response_class=HTMLResponse)
async def fill_html2_api(payload:FillHtmlPayload):
    try:
        html_content = payload.template_html
        json_text = payload.data
        # json_content = payload.data
        # json_text = json_content["extracted_json"]

        filled_html = fill_html2(html_content, json_text)
        # filled_html = fill_html_with_json(html_content, json_text)
        return HTMLResponse(content=filled_html)
    except Exception as e:
        print("Error in /fill-html:", str(e))
        return HTMLResponse(content=f"Internal Server Error: {e}", status_code=500)


@app.post("/docx-to-json")
async def docx_to_json(request: Request):
    form  = await request.form()
    file = None
    for v in form.values():
        if hasattr(v,"filename") or hasattr(v,'file'):
            file = v
            break
    if file is None:
        return{"Error" :"No file provided"}
    with tempfile.NamedTemporaryFile(delete=False) as tmp:
        tmp.write(await file.read()) # type: ignore
        tmp_path = tmp.name
    out = extract_docx_json(tmp_path)
    os.remove(tmp_path)
    return out


import time, pymupdf
@app.post("/pdf-to-text")
async def pdf_to_text(request: Request):
    """
    Upload a PDF via multipart/form-data and extract text using PyMuPDF.

    Uncomment ONE of the two extraction blocks below:
    1) LINE-BY-LINE (simple linear page text)
    2) BLOCK-ORDERED (sorted by y0 then x0 for better reading order)
    """

    # Parse multipart/form-data and find the uploaded file
    try:
        form = await request.form()
    except Exception as e:
        return JSONResponse({"Error": f"Invalid form-data: {e}"}, status_code=400)

    file = None
    for v in form.values():
        if hasattr(v, "filename") or hasattr(v, "file"):
            file = v
            break

    if file is None:
        return JSONResponse({"Error": "No file provided"}, status_code=400)

    # Read file bytes from the upload (avoid temp files)
    try:
        pdf_bytes = await file.read()
    except Exception as e:
        return JSONResponse({"Error": f"Failed to read uploaded file: {e}"}, status_code=400)

    if not pdf_bytes:
        return JSONResponse({"Error": "Uploaded file is empty"}, status_code=400)

    # Open PDF via PyMuPDF (from bytes)
    start = time.time()
    try:
        doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        return JSONResponse({"Error": f"Failed to open PDF: {e}"}, status_code=400)

    parts = []
    try:
        # ----------------------------------------------------------
        # OPTION 2: BLOCK-ORDERED (better reading order for columns)
        # ----------------------------------------------------------
        for page in doc:
            blocks = page.get_text("blocks")
            # Each block: (x0, y0, x1, y1, text, block_no, ...)
            blocks.sort(key=lambda b: (b[1], b[0]))  # sort by top, then left
            for b in blocks:
                text_block = (b[4] or "").strip()
                if text_block:
                    parts.append(text_block)

        # NOTE: Uncomment ONLY ONE of the above sections.
        # If neither is uncommented, parts will be empty.
        pass

    finally:
        doc.close()

    duration = time.time() - start
    print(f"Extraction time: {duration:.2f} seconds")

    return ("\n".join(parts))


def test_html_to_docx_bytes(html: str) -> bytes:
    soup = BeautifulSoup(html, "html.parser")
    document = Document()
    h2_counter = 0
    h3_counter = 0
    # ================= STEP 1: EXTRACT <header> =================

    header = {
        "msme": "",
        "title": "",
        "ipo": ""
    }

    header_tag = soup.find("header")

    if header_tag:
        divs = header_tag.find_all("div", class_="subtitle")
        if len(divs) >= 1:
            header["msme"] = divs[0].get_text(strip=True)
        if len(divs) >= 2:
            header["title"] = divs[1].get_text(strip=True)
        if len(divs) >= 3:
            header["ipo"] = divs[2].get_text(strip=True)

        header_tag.decompose()  # 🔥 remove header from HTML

    # ================= STEP 2: RENDER DOCX HEADER =================

    if any(header.values()):

        table = document.add_table(rows=1, cols=1)
        table.autofit = True

        left_p = table.rows[0].cells[0].paragraphs[0]
        left_p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # LEFT: MSME
        if header["msme"]:
            add_text(left_p, header["msme"], size=11, bold=True)

        document.add_paragraph("")  # spacing

        # CENTER: Title
        if header["title"]:
            p = document.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_text_with_missing_highlight(p,header["title"],11)
            

        # CENTER + UNDERLINE: IPO line
        if header["ipo"]:
            p = document.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_text_with_missing_highlight(p,header["title"],11)
            
    # Only body content
    root = soup.body if soup.body else soup

    def handle_element(el):

        # Ignore style/script/head
        if isinstance(el, Tag) and el.name in ["style", "script", "head"]:
            return

        # Plain text
        if isinstance(el, NavigableString):
            text = el.strip()
            if text:
                add_text(document, text, size=10)
            return

        # Headings → 11 pt, bold, NO style
        if el.name in ["h1", "h2", "h3"]:
            text = el.get_text(" ", strip=True)
            if text:
                p = document.add_paragraph(text)
                p.style = "List Number"
                for run in p.runs:
                    run.font.name = FONT_NAME
                    run.font.size = Pt(11)
                    run.bold = True
            return
        nonlocal h2_counter, h3_counter
        if el.name =="h2":
            text = el.get_text(" ",strip=True)
            if text:
                h2_counter +=1
                
                numbered = f"{h2_counter}. {text}"
                p = document.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                add_text(p,numbered, size=11,bold=True)
            return   

        if el.name =="h3":
            text = el.get_text(" ",strip=True)
            if text:
                h3_counter +=1
                
                numbered = f"{h2_counter}.{h3_counter} {text}"
                p = document.add_paragraph()
                p.paragraph_format.left_indent = Pt(18)
                add_text(p,numbered, size=10,bold=False)
            return 
        # Paragraphs → 10 pt (handle <br>)
        if el.name == "p":
            p = document.add_paragraph()
            p.style = None
            def add_inline(node, bold=False, italic=False):
                if isinstance(node, NavigableString):
                    txt = str(node)
                    if txt:
                        if txt.strip().lower() == "business to fill":
                            run = p.add_run(txt.strip())
                            format_run(run,10)
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        else:
                            add_text_with_missing_highlight(p,txt,10)    
                        return

                if isinstance(node,Tag):
                    p.add_run("\n")
                    return 

                if node.name in ["b","strong"]:
                    for c in node.children:
                        add_inline(c,bold=True,italic=italic)
                    return
                
                if node.name in ["i","em"]:
                    for c in node.children:
                        add_inline(c,bold=bold,italic=True)
                    return
                
                # neseted tags
                for c in node.children:
                    add_inline(c,bold=True,italic=italic)
                    return           
                
            for child in el.children:
                add_inline(child)   
            
            return
        
        # Handle div
        if el.name == "div":
            p = document.add_paragraph()
            p.style = None

            for node in el.children:
                if isinstance(node,NavigableString):
                    txt = str(node)
                    if txt:
                        if txt.strip().lower() == "business to fill":
                            run = p.add_run(txt.strip())
                            format_run(run,10)
                            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                        else:
                            add_text_with_missing_highlight(p,txt,10)
                else:
                    handle_element(node) # Nested tags

            return                
       
        # Unordered list
        if el.name == "ul":
            for li in el.find_all("li", recursive=False):
                text = li.get_text(" ", strip=True)
                if text:
                    p = document.add_paragraph(text, style="List Bullet")

                    # enforce font
                    for run in p.runs:
                        run.font.name = FONT_NAME
                        run.font.size = Pt(10)
            return

        # Ordered list
        if el.name == "ol":
            for idx, li in enumerate(el.find_all("li", recursive=False), start=1):
                text = li.get_text(" ", strip=True)
                if text:
                    p = document.add_paragraph()
                    p.style = None
                    add_text(p, f"{idx}. {text}", size=10)
            return
        if el.name == "table":

            rows = el.find_all("tr")
            if not rows:
                return

            parsed_rows = []
            max_cols = 0

            # 🔥 calculate REAL columns considering colspan
            for r in rows:
                cells = r.find_all(["th", "td"], recursive=False)
                if not cells:
                    continue

                parsed_rows.append(cells)

                col_count = 0
                for c in cells:
                    col_count += int(c.get("colspan", 1))

                max_cols = max(max_cols, col_count)


            # create table
            table = document.add_table(rows=len(parsed_rows), cols=max_cols)
            table.autofit = False
            table.style = None   # ❌ remove ugly Word grid


            # 🔥 fill rows
            for r_idx, cells in enumerate(parsed_rows):

                col_ptr = 0

                for cell_el in cells:

                    colspan = int(cell_el.get("colspan", 1))

                    cell = table.rows[r_idx].cells[col_ptr]
                    set_cell_border(cell)

                    # ===== MERGE (handles Pre-IPO / Post-IPO headers) =====
                    if colspan > 1:
                        merge_to = table.rows[r_idx].cells[col_ptr + colspan - 1]
                        cell = cell.merge(merge_to)

                    # clear
                    cell.text = ""

                    p = cell.paragraphs[0]
                    p.style = None

                    # padding for clean look
                    p.paragraph_format.space_before = Pt(3)
                    p.paragraph_format.space_after = Pt(3)

                    cell_text = cell_el.get_text(" ", strip=True)
                    if cell_text.strip().lower() == "business to fill":
                        run = p.add_run(cell_text.strip())
                        format_run(run,10)
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    else:
                        add_text_with_missing_highlight(p,cell_text.strip(),10)
                    col_ptr += colspan

            return

        # Containers
        for child in el.children:
            handle_element(child)

    for child in root.children:
        handle_element(child)

    buf = BytesIO()
    document.save(buf)
    buf.seek(0)
    return buf.getvalue()


def html_to_docx_bytes(html: str) -> bytes:
    """
    Convert HTML (like your IM template) to a reasonably structured DOCX.
    - h1/h2/h3 -> Word headings
    - p       -> paragraphs
    - ul/ol   -> bullet/numbered lists
    - table   -> grid table with header row
    """
    soup = BeautifulSoup(html, "html.parser")
    document = Document()

    # Optional: add main title from <title> if present
    if soup.title and soup.title.string:
        title_para = document.add_paragraph(soup.title.string.strip())
        title_para.style = "Title"
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Work on <body> if present, else whole soup
    root: Tag = soup.body if soup.body else soup

    def handle_element(el: Tag, doc: Document):
        # Skip non-tag nodes handled via parent
        if isinstance(el, NavigableString):
            text = el.strip()
            if text:
                doc.add_paragraph(text)
            return

        # Headings
        if el.name in ["h1", "h2", "h3"]:
            text = el.get_text(strip=True)
            if not text:
                return
            level = {"h1": 1, "h2": 2, "h3": 3}[el.name]
            doc.add_heading(text, level=level)
            return

        # Paragraphs
        if el.name == "p":
            text = el.get_text(" ", strip=True)
            if text:
                doc.add_paragraph(text)
            return

        # Unordered list (bullets)
        if el.name == "ul":
            for li in el.find_all("li", recursive=False):
                text = li.get_text(" ", strip=True)
                if text:
                    p = doc.add_paragraph(text)
                    p.style = "List Bullet"
            return

        # Ordered list (numbers)
        if el.name == "ol":
            for li in el.find_all("li", recursive=False):
                text = li.get_text(" ", strip=True)
                if text:
                    p = doc.add_paragraph(text)
                    p.style = "List Number"
            return

        # Tables
        if el.name == "table":
            rows = el.find_all("tr")
            if not rows:
                return

            # Determine max columns from first row with cells
            max_cols = 0
            row_cells = []
            for r in rows:
                cells = r.find_all(["th", "td"], recursive=False)
                if cells:
                    row_cells.append(cells)
                    max_cols = max(max_cols, len(cells))

            if max_cols == 0:
                return

            table = doc.add_table(rows=len(row_cells), cols=max_cols)
            table.style = "Table Grid"

            for r_idx, cells in enumerate(row_cells):
                row = table.rows[r_idx]
                for c_idx in range(max_cols):
                    cell = row.cells[c_idx]
                    if c_idx < len(cells):
                        text = cells[c_idx].get_text(" ", strip=True)
                        cell.text = text

            return

        # Structural wrappers: handle children
        if el.name in ["div", "section", "main", "header", "footer"]:
            for child in el.children:
                if isinstance(child, (Tag, NavigableString)):
                    handle_element(child, doc)
            return

        # Fallback: just process children
        for child in el.children:
            if isinstance(child, (Tag, NavigableString)):
                handle_element(child, doc)

    # Process top-level children of root
    for child in root.children:
        if isinstance(child, (Tag, NavigableString)):
            handle_element(child, document)

    # Save to bytes
    buf = BytesIO()
    document.save(buf)
    buf.seek(0)
    return buf.getvalue()


@app.post('/html-to-docx')
async def html_to_docx(
    request: Request,
    html: Optional[str] = Form(None),
    file: Optional[UploadFile] = Form(None)
):
    try:
        file_name = "output.docx"
        if request.headers.get("content-type","").startswith("application/json"):
            payload = HtmlToDocxPayload(**await request.json())
            html_content = payload.html
            file_name = payload.file_name
        elif file:
            html_content = (await file.read()).decode("utf-8")
        elif html:
            html_content = html
        else:
            html_content = (await request.body()).decode("utf-8")

        if not html_content.strip():
            return JSONResponse(
                status_code = 400,
                content = {"error":"HTML content is empty"},
            )
        docx_bytes = test_html_to_docx_bytes(html_content)

        return StreamingResponse(
            BytesIO(docx_bytes),
            media_type= "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition":f'attachment;filename="{file_name}"'
            },
        )


            
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error":str(e)}
        )







