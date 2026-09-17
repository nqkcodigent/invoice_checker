"""
invoice_checker.py — Đối chiếu hóa đơn điện tử (PDF) với bảng tổng hợp chi phí (Excel).
"""

from __future__ import annotations

import argparse
import html
import io
import json
import os
import re
import sys
import unicodedata
import webbrowser
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

import pandas as pd
import pdfplumber
from tabulate import tabulate

# ============================== CẤU HÌNH ==============================
EXCEL_FILE = "SIM_FM-ACC-03-4 Expense Claim_NguyenHuyenNgan_August.2026.xlsx"
SHEET_NAME = "Claim"                    
PDF_FOLDER = "August.26/August.26"      
HEADER_KEYWORD = "số chứng từ"          
EXPECTED_COMPANY = "CÔNG TY TNHH ACCLIME OUTSOURCING"
EXPECTED_TAXCODE = "0316791220"
OUTPUT_HTML = "report.html"
# =======================================================================


# ------------------------------- MÔ HÌNH DỮ LIỆU -------------------------------

@dataclass
class ExcelRow:
    stt: str
    no: str                 
    no_norm: str            
    date: str               
    amount: str             
    description: str = ""
    notes: str = ""

@dataclass
class PdfData:
    path: str
    no: str = ""
    no_norm: str = ""
    date: str = ""
    amount: str = ""
    company: str = ""
    tax: str = ""
    has_signature: bool = False
    error: str = ""

@dataclass
class Result:
    pdf: PdfData | None
    excel: ExcelRow | None
    fields: dict[str, tuple[str, str, bool]] = field(default_factory=dict)
    status: str = ""        
    status_label: str = ""


# ------------------------------- TIỆN ÍCH CHUẨN HÓA -------------------------------

def norm_number(value: Any) -> str:
    digits = re.sub(r"[^\d]", "", str(value))
    return str(int(digits)) if digits else ""

def norm_amount(value: Any) -> str:
    s = str(value).strip()
    if s.endswith(".0"):          
        s = s[:-2]
    return s.replace(".", "").replace(",", "")

def norm_date(value: Any) -> str:
    if isinstance(value, (datetime, pd.Timestamp)):
        return pd.Timestamp(value).strftime("%d.%m.%Y")
    s = str(value).strip()
    m = re.search(r"(\d{1,2})[\./](\d{1,2})[\./](\d{2,4})", s)
    if m:
        d, mo, y = m.groups()
        y = y if len(y) == 4 else f"20{y}"
        return f"{int(d):02d}.{int(mo):02d}.{y}"
    return s


# ------------------------------- ĐỌC FILE EXCEL -------------------------------

def _find_col(df: pd.DataFrame, *keywords: str) -> str | None:
    for col in df.columns:
        low = str(col).lower()
        if any(k.lower() in low for k in keywords):
            return col
    return None

def load_excel_rows(excel_path: str, sheet_name: str) -> list[ExcelRow]:
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Không tìm thấy file Excel: {excel_path}")

    raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    header_idx = next(
        (i for i, row in raw.iterrows() if HEADER_KEYWORD in " ".join(str(c) for c in row.values).lower()),
        -1,
    )
    if header_idx == -1:
        raise ValueError(f"Không tìm thấy dòng tiêu đề (chứa '{HEADER_KEYWORD}') trong sheet '{sheet_name}'.")

    df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=header_idx)
    df.columns = df.columns.astype(str).str.replace("\n", " ").str.strip()

    col_no = _find_col(df, "số chứng từ", "documents no")
    col_date = _find_col(df, "ngày chứng từ", "document date")
    col_amount = _find_col(df, "thành tiền", "amount")
    col_stt = _find_col(df, "stt")
    col_desc = _find_col(df, "diễn giải", "description")
    col_notes = _find_col(df, "ghi chú", "notes")

    rows: list[ExcelRow] = []
    for _, r in df.iterrows():
        no = r.get(col_no) if col_no else None
        if pd.isna(no): continue
        no_str = str(no).strip()
        no_norm = norm_number(no_str)
        if not no_norm: continue
        rows.append(
            ExcelRow(
                stt=str(r.get(col_stt, "")).strip() if col_stt else "",
                no=no_str,
                no_norm=no_norm,
                date=norm_date(r.get(col_date, "")) if col_date else "",
                amount=norm_amount(r.get(col_amount, "")) if col_amount else "",
                description=str(r.get(col_desc, "")).strip() if col_desc else "",
                notes=str(r.get(col_notes, "")).strip() if col_notes else "",
            )
        )
    return rows


# ------------------------------- TRÍCH XUẤT DỮ LIỆU PDF -------------------------------

def extract_pdf_text(pdf_path: str) -> str:
    """Trích toàn bộ văn bản từ PDF và CỰC KỲ QUAN TRỌNG: Chuẩn hóa font Unicode (NFC)"""
    parts: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                # Xử lý dứt điểm lỗi máy tính đọc tiếng Việt từ PDF bị tách rấu (NFD -> NFC)
                # Nếu không có dòng này, regex tìm chữ ký sẽ luôn báo "Không"
                parts.append(unicodedata.normalize("NFC", text))
    return "\n".join(parts)


def extract_invoice_no(text: str, filename: str) -> str:
    """Xử lý định dạng lồng chữ: 'Số hóa đơn (Invoice No): 597518'"""
    m = re.search(r"Số[^\(]*\((?:Invoice\s*)?No\.?\)\s*:?\s*(\d+)", text, re.IGNORECASE)
    if m: return m.group(1)

    m = re.search(r"Số(?: hóa đơn)?\s*:\s*(\d+)", text, re.IGNORECASE)
    if m: return m.group(1)

    m = re.search(r"(\d{3,})\s*\n\s*Số\s*\(", text, re.IGNORECASE)
    if m: return m.group(1)

    m = re.search(r"Inv[_\- ]?0*(\d+)", filename, re.IGNORECASE)
    return m.group(1) if m else ""


AMOUNT_PATTERNS: list[tuple[re.Pattern[str], bool]] = [
    # Vietnam Airlines: "Tổng số tiền thanh toán (Grand Total...):\n | | 5.899.000"
    (re.compile(r"Tổng số tiền thanh toán\s*(?:\([^)]*\))?\s*:?[\s\|]*([\d\.]+(?:[\s\|]+[\d\.]+)*)", re.IGNORECASE), True),
    (re.compile(r"Tổng cộng số tiền đã có thuế GTGT\s*:?\s*\n?\s*([\d\.]+)", re.IGNORECASE), False),
    (re.compile(r"Tổng cộng hóa đơn\s*(?:\([^)]*\))?\s*:?\s*([\d\.]+)", re.IGNORECASE), False),
    (re.compile(r"Tổng cộng tiền thanh toán\s*(?:\([^)]*\))?\s*:?\s*([\d\.]+(?:[ \t]+[\d\.]+)*)", re.IGNORECASE), True),
    (re.compile(r"Tổng tiền thanh toán\s*(?:\([^)]*\))?\s*:?\s*([\d\.]+)", re.IGNORECASE), False),
    (re.compile(r"Tổng cộng\s*(?:\([^)]*\))?\s*:?\s*([\d\.]+(?:[ \t]+[\d\.]+)*)"), True),
]

def extract_amount(text: str) -> str:
    """
    Chiến thuật Mới: Tìm khối text ngay sau từ khóa (khoảng 80 ký tự) và quét tất cả các số bên trong.
    Khắc phục triệt để lỗi PDF xuống dòng, chèn ký tự bảng (|), hoặc khoảng trắng nhiễu.
    """
    PATTERNS = [
        (re.compile(r"Tổng số tiền thanh toán(.{1,80})", re.IGNORECASE | re.DOTALL), True),
        (re.compile(r"Tổng cộng tiền thanh toán(.{1,80})", re.IGNORECASE | re.DOTALL), True),
        (re.compile(r"Tổng cộng số tiền đã có thuế GTGT(.{1,80})", re.IGNORECASE | re.DOTALL), False),
        (re.compile(r"Tổng cộng hóa đơn(.{1,80})", re.IGNORECASE | re.DOTALL), False),
        (re.compile(r"Tổng tiền thanh toán(.{1,80})", re.IGNORECASE | re.DOTALL), False),
        (re.compile(r"Tổng cộng([^a-z]{1,40})", re.IGNORECASE | re.DOTALL), True),
    ]
    
    for pattern, take_last in PATTERNS:
        m = pattern.search(text)
        if m:
            block = m.group(1)
            # Lọc trong vòng 80 ký tự đó, rút ra tất cả các cụm số có dạng tiền (vd: 5.899.000) hoặc số liền
            numbers = re.findall(r"\d{1,3}(?:\.\d{3})+|\d+", block)
            if numbers:
                return numbers[-1] if take_last else numbers[0]
                
    return ""

DATE_RE = re.compile(
    r"Ngày\s*(?:\([^)]*\))?\s*(\d{1,2})\s*tháng\s*(?:\([^)]*\))?\s*(\d{1,2})\s*năm\s*(?:\([^)]*\))?\s*(\d{4})",
    re.IGNORECASE,
)

def extract_date(text: str) -> str:
    m = DATE_RE.search(text)
    if m:
        d, mo, y = m.groups()
        return f"{int(d):02d}.{int(mo):02d}.{y}"
    return ""


def extract_company(text: str) -> str:
    # Quét nhãn 'Tên người mua (Buyer)' thay vì chỉ 'Tên đơn vị'
    m = re.search(r"(?:Tên đơn vị|Tên người mua)\s*(?:\([^)]*\))?\s*:\s*(.+)", text, re.IGNORECASE)
    if m: return m.group(1).strip()
    if "ACCLIME" in text.upper(): return EXPECTED_COMPANY
    return ""


def extract_taxcode(text: str) -> str:
    """Quét toàn bộ mã số thuế và ưu tiên lấy mã khớp cấu hình EXPECTED_TAXCODE"""
    codes = [
        re.sub(r"[^\d]", "", m.group(1))
        # Cho phép nhãn "MST", quét không phân biệt hoa/thường
        for m in re.finditer(r"(?:Mã số thuế|MST)\s*(?:\([^)]*\))?\s*:\s*([\d\s\-]+)", text, re.IGNORECASE)
    ]
    codes = [c for c in codes if c]
    if not codes:
        codes = re.findall(r"\b\d{10}\b", text)
    if not codes:
        return ""
    
    # Lấy đúng mã ACCLIME (0316791220) trong số các mã tìm được
    for c in codes:
        if c == EXPECTED_TAXCODE: return c
    return codes[-1]


SIGNATURE_RE = re.compile(r"(Signature Valid|Ký bởi|ký điện tử|Signed by|đã ký)", re.IGNORECASE)

def extract_signature(text: str) -> bool:
    """
    Chiến thuật Mới: San phẳng mọi chướng ngại vật từ PDF.
    Xóa sạch khoảng trắng, ký tự ẩn, dấu tiếng Việt để chống lại lỗi font bị rời rạc (vd: 'K ý  đ i ệ n').
    """
    # 1. Thử bắt bằng regex chuẩn
    if re.search(r"(Signature Valid|Ký bởi|ký điện tử|Signed by|đã ký|Signed date|Ký ngày)", text, re.IGNORECASE):
        return True
    
    # 2. Xử lý hạng nặng: Cạo sạch khoảng trắng, xuống dòng, và các ký tự ẩn tàng hình
    text_clean = re.sub(r"[\s\x00-\x1f]", "", text.lower())
    
    # Bỏ toàn bộ dấu Tiếng Việt (NFD)
    text_clean = "".join(c for c in unicodedata.normalize("NFD", text_clean) if unicodedata.category(c) != "Mn")
    
    # Đổi ký tự đặc biệt 'đ' thành 'd'
    text_clean = text_clean.replace("đ", "d")
    
    # Từ khóa chữ ký bây giờ được so sánh trên chuỗi tiếng Việt không dấu, không khoảng trắng
    keywords = [
        "signaturevalid", "kyboi", "kydientu", "signedby", 
        "dakydientu", "kyngay", "signeddate", "daduocky"
    ]
    
    return any(k in text_clean for k in keywords)


def parse_pdf(pdf_path: str) -> PdfData:
    data = PdfData(path=pdf_path)
    try:
        text = extract_pdf_text(pdf_path)
    except Exception as exc: 
        data.error = f"Lỗi đọc file: {exc}"
        return data

    data.no = extract_invoice_no(text, os.path.basename(pdf_path))
    data.no_norm = norm_number(data.no)
    data.date = extract_date(text)
    data.amount = norm_amount(extract_amount(text))
    data.company = extract_company(text)
    data.tax = extract_taxcode(text)
    data.has_signature = extract_signature(text)
    return data


# ------------------------------- KHỚP & SO SÁNH -------------------------------

def match_excel_row(excel_rows: list[ExcelRow], pdf_no_norm: str) -> ExcelRow | None:
    if not pdf_no_norm: return None
    return next((r for r in excel_rows if r.no_norm == pdf_no_norm), None)

def _fold_diacritics(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", s.lower()) if unicodedata.category(c) != "Mn")

def compare(excel: ExcelRow, pdf: PdfData) -> dict[str, tuple[str, str, bool]]:
    no_ok = bool(excel.no_norm) and excel.no_norm == pdf.no_norm
    date_ok = bool(excel.date) and excel.date == pdf.date
    amount_ok = bool(excel.amount) and excel.amount == pdf.amount
    company_ok = bool(pdf.company) and _fold_diacritics(EXPECTED_COMPANY) in _fold_diacritics(pdf.company)
    tax_ok = bool(pdf.tax) and pdf.tax == EXPECTED_TAXCODE
    return {
        "so": (excel.no, pdf.no, no_ok),
        "date": (excel.date, pdf.date, date_ok),
        "amount": (excel.amount, pdf.amount, amount_ok),
        "company": (EXPECTED_COMPANY, pdf.company, company_ok),
        "tax": (EXPECTED_TAXCODE, pdf.tax, tax_ok),
        "signature": ("Có chữ ký", "Có" if pdf.has_signature else "Không", pdf.has_signature),
    }

def build_result(pdf: PdfData, excel: ExcelRow | None) -> Result:
    if pdf.error: return Result(pdf=pdf, excel=excel, status="error", status_label="Lỗi đọc file")
    if excel is None: return Result(pdf=pdf, excel=None, status="no_excel", status_label="Không có trong Excel")
    fields = compare(excel, pdf)
    all_ok = all(ok for _, _, ok in fields.values())
    status = "ok" if all_ok else "mismatch"
    label = "✅ Khớp" if all_ok else "❌ Lệch"
    return Result(pdf=pdf, excel=excel, fields=fields, status=status, status_label=label)


# ------------------------------- XUẤT KẾT QUẢ -------------------------------

def export_csv(results: list[Result], excel_missing: list[ExcelRow]) -> str:
    import csv
    header = ["STT", "File PDF", "Số CT (Excel)", "Số CT (PDF)", "Ngày (Excel)", "Ngày (PDF)",
              "Thành tiền (Excel)", "Thành tiền (PDF)", "Tên công ty", "Mã số thuế", "Chữ ký", "Kết luận", "Ghi chú"]
    rows: list[list[str]] = []
    for r in results:
        if r.pdf is None: continue
        pdf = r.pdf
        stt = r.excel.stt if r.excel else ""
        no_ex = r.excel.no if r.excel else ""
        no_pdf = pdf.no
        date_ex = r.excel.date if r.excel else ""
        date_pdf = pdf.date
        amt_ex = r.excel.amount if r.excel else ""
        amt_pdf = pdf.amount
        company = pdf.company if pdf.company else (r.excel.notes if r.excel else "")
        tax = pdf.tax
        sign = "Có" if pdf.has_signature else ("Không" if pdf.error else "Không tìm thấy")
        desc = (r.excel.notes or r.excel.description) if r.excel else ""
        rows.append([stt, os.path.basename(pdf.path), no_ex, no_pdf, date_ex, date_pdf,
                     amt_ex, amt_pdf, company, tax, sign, r.status_label, desc])
    for row in excel_missing:
        rows.append([row.stt, "", row.no, "", row.date, "", row.amount, "",
                     "", "", "", "Thiếu file PDF", row.notes or row.description])

    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(header)
    writer.writerows(rows)
    return "\ufeff" + buf.getvalue()


def render_console(results: list[Result], excel_missing: list[ExcelRow]) -> None:
    matched = sum(1 for r in results if r.status == "ok")
    mismatch = sum(1 for r in results if r.status == "mismatch")
    no_excel = sum(1 for r in results if r.status == "no_excel")
    errors = sum(1 for r in results if r.status == "error")

    print("\n" + "=" * 100)
    print(f"📊 KẾT QUẢ ĐỐI CHIẾU HÓA ĐƠN PDF ↔ EXCEL (sheet '{SHEET_NAME}')")
    print("=" * 100)
    print(f"   Tổng PDF: {len(results)} | ✅ Khớp: {matched} | ❌ Lệch: {mismatch} | "
          f"⚠️ Không có trong Excel: {no_excel} | 🔴 Lỗi đọc: {errors} | "
          f"📌 Dòng Excel thiếu PDF: {len(excel_missing)}")

    headers = ["STT", "File PDF", "Số CT (Excel)", "Số CT (PDF)", "Ngày", "Thành tiền", "MST", "Chữ ký", "Kết luận"]
    table: list[list[str]] = []
    for r in results:
        if r.pdf is None: continue
        stt = r.excel.stt if r.excel else "—"
        no_ex = r.excel.no if r.excel else "—"
        table.append([stt, os.path.basename(r.pdf.path), no_ex, r.pdf.no, r.pdf.date, r.pdf.amount or "—", r.pdf.tax or "—", "✅" if r.pdf.has_signature else "❌", r.status_label])
    print(tabulate(table, headers=headers, tablefmt="fancy_grid"))


def render_html(results: list[Result], excel_missing: list[ExcelRow]) -> str:
    rows_json: list[dict[str, Any]] = []

    for r in results:
        if r.pdf is None: continue
        pdf = r.pdf
        base = {
            "stt": r.excel.stt if r.excel else "—",
            "file": os.path.basename(pdf.path),
            "desc": (r.excel.notes or r.excel.description) if r.excel else "",
            "status": r.status,
            "status_label": r.status_label,
        }
        if r.status == "no_excel":
            base.update(no_excel="—", no_pdf=pdf.no, date_excel="—", date_pdf=pdf.date,
                        amt_excel="—", amt_pdf=pdf.amount, company=pdf.company, tax=pdf.tax,
                        sign="✅" if pdf.has_signature else "❌")
            rows_json.append(base)
            continue
        if r.status == "error":
            base.update(no_excel="—", no_pdf="", date_excel="—", date_pdf="", amt_excel="—",
                        amt_pdf="", company="", tax="", sign="", error=pdf.error)
            rows_json.append(base)
            continue
        f = r.fields
        base.update(
            no_excel=f["so"][0], no_pdf=f["so"][1], no_ok=f["so"][2],
            date_excel=f["date"][0], date_pdf=f["date"][1], date_ok=f["date"][2],
            amt_excel=f["amount"][0], amt_pdf=f["amount"][1], amt_ok=f["amount"][2],
            company=f["company"][1], company_ok=f["company"][2],
            tax=f["tax"][1], tax_ok=f["tax"][2],
            sign="✅" if f["signature"][2] else "❌",
        )
        rows_json.append(base)

    for row in excel_missing:
        rows_json.append({
            "stt": row.stt, "file": "—", "desc": row.notes or row.description,
            "status": "no_pdf", "status_label": "Thiếu file PDF",
            "no_excel": row.no, "no_pdf": "—", "date_excel": row.date, "date_pdf": "—",
            "amt_excel": row.amount, "amt_pdf": "—", "company": "—", "tax": "—", "sign": "—",
        })

    counts = {
        "total": len(results), "ok": sum(1 for r in results if r.status == "ok"),
        "mismatch": sum(1 for r in results if r.status == "mismatch"),
        "no_excel": sum(1 for r in results if r.status == "no_excel"),
        "errors": sum(1 for r in results if r.status == "error"),
        "no_pdf": len(excel_missing),
    }

    payload = json.dumps(rows_json, ensure_ascii=False)

    html_doc = f"""<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Đối chiếu hóa đơn PDF ↔ Excel</title>
<style>
  :root {{ --ok:#e6f4ea; --mismatch:#fdecea; --warn:#fff8e1; --border:#d0d7de; }}
  * {{ box-sizing:border-box; }}
  body {{ font-family:"Segoe UI",Arial,sans-serif; margin:0; background:#f6f8fa; color:#1f2328; }}
  header {{ background:#24292f; color:#fff; padding:18px 28px; }}
  header h1 {{ margin:0 0 10px; font-size:20px; }}
  .cards {{ display:flex; gap:12px; flex-wrap:wrap; }}
  .card {{ background:rgba(255,255,255,.12); border-radius:8px; padding:8px 16px; font-size:13px; }}
  .card b {{ font-size:22px; display:block; }}
  .toolbar {{ padding:14px 28px; display:flex; gap:12px; align-items:center; background:#fff; border-bottom:1px solid var(--border); position:sticky; top:0; z-index:5; }}
  .toolbar input {{ flex:1; max-width:380px; padding:8px 12px; border:1px solid var(--border); border-radius:6px; font-size:14px; }}
  .toolbar select {{ padding:8px 12px; border:1px solid var(--border); border-radius:6px; font-size:14px; }}
  .wrap {{ padding:22px 28px 40px; }}
  table {{ width:100%; border-collapse:collapse; background:#fff; box-shadow:0 1px 3px rgba(0,0,0,.08); }}
  th, td {{ border-bottom:1px solid var(--border); padding:8px 10px; font-size:13px; text-align:left; vertical-align:top; }}
  th {{ background:#f0f2f5; cursor:pointer; user-select:none; white-space:nowrap; position:sticky; top:57px; }}
  th:hover {{ background:#e3e6ea; }}
  tr.ok td {{ background:var(--ok); }}
  tr.mismatch td {{ background:var(--mismatch); }}
  tr.no_excel td, tr.no_pdf td {{ background:var(--warn); }}
  tr.error td {{ background:#f3d9d9; }}
  .badge {{ display:inline-block; padding:2px 10px; border-radius:12px; font-size:12px; font-weight:600; }}
  .b-ok {{ background:#1a7f37; color:#fff; }}
  .b-mismatch {{ background:#cf222e; color:#fff; }}
  .b-no_excel, .b-no_pdf {{ background:#9a6700; color:#fff; }}
  .b-error {{ background:#8250df; color:#fff; }}
  .mono {{ font-variant-numeric:tabular-nums; }}
  .desc {{ color:#57606a; font-size:12px; max-width:220px; }}
  .file {{ font-size:12px; color:#57606a; }}
  .empty {{ text-align:center; color:#8c959f; padding:30px; }}
</style>
</head>
<body>
<header>
  <h1>📊 Đối chiếu hóa đơn PDF ↔ Excel</h1>
  <div class="cards">
    <div class="card"><b>{counts['total']}</b>Tổng PDF</div>
    <div class="card"><b>{counts['ok']}</b>✅ Khớp</div>
    <div class="card"><b>{counts['mismatch']}</b>❌ Lệch</div>
    <div class="card"><b>{counts['no_excel']}</b>⚠️ Không có trong Excel</div>
    <div class="card"><b>{counts['no_pdf']}</b>📌 Excel thiếu PDF</div>
    <div class="card"><b>{counts['errors']}</b>🔴 Lỗi đọc</div>
  </div>
</header>
<div class="toolbar">
  <input id="search" type="text" placeholder="🔍 Tìm theo số chứng từ, tên file, ghi chú...">
  <select id="statusFilter">
    <option value="">Tất cả trạng thái</option>
    <option value="ok">✅ Khớp</option>
    <option value="mismatch">❌ Lệch</option>
    <option value="no_excel">⚠️ Không có trong Excel</option>
    <option value="no_pdf">📌 Thiếu file PDF</option>
    <option value="error">🔴 Lỗi đọc</option>
  </select>
</div>
<div class="wrap">
<table id="report">
  <thead><tr>
    <th data-key="stt">STT</th>
    <th data-key="file">File PDF</th>
    <th data-key="no_excel">Số chứng từ (Excel)</th>
    <th data-key="no_pdf">Số chứng từ (PDF)</th>
    <th data-key="date_excel">Ngày lập</th>
    <th data-key="amt_excel">Thành tiền</th>
    <th data-key="company">Tên công ty</th>
    <th data-key="tax">Mã số thuế</th>
    <th data-key="sign">Chữ ký</th>
    <th data-key="status">Kết luận</th>
    <th>Ghi chú</th>
  </tr></thead>
  <tbody></tbody>
</table>
</div>
<script>
const DATA = {payload};
let sortKey = "stt", sortDir = 1;

function cellText(r, k) {{
  if (k === "sign") return r.sign || "—";
  if (k === "status") return r.status_label;
  return r[k] !== undefined && r[k] !== "" ? String(r[k]) : "—";
}}

function render() {{
  const q = document.getElementById("search").value.trim().toLowerCase();
  const st = document.getElementById("statusFilter").value;
  let rows = DATA.filter(r => (st === "" || r.status === st));
  if (q) rows = rows.filter(r => Object.values(r).join(" ").toLowerCase().includes(q));
  rows.sort((a, b) => {{
    let va = cellText(a, sortKey), vb = cellText(b, sortKey);
    if (["no_excel","no_pdf","amt_excel","amt_pdf"].includes(sortKey)) {{
      va = parseFloat(String(va).replace(/[^0-9.]/g,"")) || 0;
      vb = parseFloat(String(vb).replace(/[^0-9.]/g,"")) || 0;
      return (va - vb) * sortDir;
    }}
    return va.localeCompare(vb, "vi") * sortDir;
  }});
  const tbody = document.querySelector("#report tbody");
  tbody.innerHTML = rows.map(r => {{
    const badge = `<span class="badge b-${{r.status}}">${{r.status_label}}</span>`;
    const ok = (k) => r[k + "_ok"] === true ? "✓" : "✗";
    return `<tr class="${{r.status}}">
      <td>${{cellText(r,"stt")}}</td>
      <td class="file">${{cellText(r,"file")}}</td>
      <td class="mono">${{cellText(r,"no_excel")}}</td>
      <td class="mono">${{cellText(r,"no_pdf")}} <span style="color:${{r.no_ok ? "#1a7f37" : "#cf222e"}}">${{ok("no")}}</span></td>
      <td class="mono">${{cellText(r,"date_excel")}} → ${{cellText(r,"date_pdf")}} <span style="color:${{r.date_ok ? "#1a7f37" : "#cf222e"}}">${{ok("date")}}</span></td>
      <td class="mono">${{cellText(r,"amt_excel")}} → ${{cellText(r,"amt_pdf")}} <span style="color:${{r.amt_ok ? "#1a7f37" : "#cf222e"}}">${{ok("amt")}}</span></td>
      <td>${{cellText(r,"company")}} <span style="color:${{r.company_ok ? "#1a7f37" : "#cf222e"}}">${{ok("company")}}</span></td>
      <td class="mono">${{cellText(r,"tax")}} <span style="color:${{r.tax_ok ? "#1a7f37" : "#cf222e"}}">${{ok("tax")}}</span></td>
      <td>${{cellText(r,"sign")}}</td>
      <td>${{badge}}</td>
      <td class="desc">${{r.desc || "—"}}</td>
    </tr>`;
  }}).join("");
  if (!rows.length) tbody.innerHTML = '<tr><td colspan="11" class="empty">Không có dữ liệu phù hợp</td></tr>';
}}

document.getElementById("search").addEventListener("input", render);
document.getElementById("statusFilter").addEventListener("change", render);
document.querySelectorAll("th").forEach(th => th.addEventListener("click", () => {{
  const k = th.dataset.key;
  if (k) {{ if (sortKey === k) sortDir *= -1; else {{ sortKey = k; sortDir = 1; }} render(); }}
}}));
render();
</script>
</body>
</html>"""

    return html_doc


# ------------------------------- MAIN -------------------------------

def find_pdfs(folder: str) -> list[str]:
    pdfs: list[str] = []
    for root, _dirs, files in os.walk(folder):
        for f in sorted(files):
            if f.lower().endswith(".pdf") and not f.startswith("~$"):
                pdfs.append(os.path.join(root, f))
    return sorted(pdfs)

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", default=EXCEL_FILE)
    parser.add_argument("--sheet", default=SHEET_NAME)
    parser.add_argument("--pdf-folder", default=PDF_FOLDER)
    parser.add_argument("--output", default=OUTPUT_HTML)
    parser.add_argument("--csv", default=None)
    parser.add_argument("--no-open", action="store_true")
    return parser.parse_args()

def main() -> None:
    args = parse_args()
    excel_rows = load_excel_rows(args.excel, args.sheet)
    pdf_files = find_pdfs(args.pdf_folder)
    
    results: list[Result] = []
    matched_nos: set[str] = set()

    for pdf_path in pdf_files:
        pdf = parse_pdf(pdf_path)
        excel = match_excel_row(excel_rows, pdf.no_norm) if not pdf.error else None
        if excel is not None: matched_nos.add(excel.no_norm)
        results.append(build_result(pdf, excel))

    excel_missing = [r for r in excel_rows if r.no_norm not in matched_nos]
    render_console(results, excel_missing)

    with open(args.output, "w", encoding="utf-8") as f:
        f.write(render_html(results, excel_missing))

    if args.csv:
        with open(args.csv, "w", encoding="utf-8", newline="") as f:
            f.write(export_csv(results, excel_missing))

    if not args.no_open and sys.platform.startswith("win"):
        try: webbrowser.open(os.path.abspath(args.output).replace("\\", "/"))
        except Exception: pass

if __name__ == "__main__":
    main()
