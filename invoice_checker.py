"""
invoice_checker.py — Đối chiếu hóa đơn điện tử (PDF) với bảng tổng hợp chi phí (Excel).

Luồng xử lý:
    1. Đọc sheet 'Claim' trong file Excel, tự động tìm dòng tiêu đề bảng.
    2. Quét toàn bộ file PDF (đệ quy) trong thư mục PDF_FOLDER.
    3. Trích xuất từ PDF: Số chứng từ, ngày lập, thành tiền, tên công ty,
       mã số thuế, chữ ký (hỗ trợ nhiều định dạng: Grab, Nasco, DHL, Koi,
       Starbucks, Takahiro, Xanh SM, SACO, BICOM, Green Fast, Petro...).
    4. Khớp từng PDF với dòng Excel theo Số chứng từ (bỏ số 0 ở đầu).
    5. So sánh từng tiêu chí, xuất kết quả ra console và file HTML
       (report.html) để rà soát bằng tay.

Cách chạy:
    python invoice_checker.py
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
SHEET_NAME = "Claim"                    # sheet chứa bảng tổng hợp chi phí
PDF_FOLDER = "August.26/August.26"      # thư mục chứa hóa đơn PDF (đệ quy)
HEADER_KEYWORD = "số chứng từ"          # từ khóa để dò dòng tiêu đề trong Excel
EXPECTED_COMPANY = "CÔNG TY TNHH ACCLIME OUTSOURCING"
EXPECTED_TAXCODE = "0316791220"
OUTPUT_HTML = "report.html"
# =======================================================================


# ------------------------------- MÔ HÌNH DỮ LIỆU -------------------------------

@dataclass
class ExcelRow:
    """Một dòng chi phí trong bảng tổng hợp Excel."""
    stt: str
    no: str                 # Số chứng từ (hiển thị)
    no_norm: str            # Số chứng từ đã chuẩn hóa (bỏ số 0 ở đầu)
    date: str               # Ngày chứng từ dạng DD.MM.YYYY
    amount: str             # Thành tiền (đã làm sạch, dạng số thuần)
    description: str = ""
    notes: str = ""


@dataclass
class PdfData:
    """Dữ liệu trích xuất từ một file PDF."""
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
    """Kết quả đối chiếu của một PDF (hoặc một dòng Excel thiếu PDF)."""
    pdf: PdfData | None
    excel: ExcelRow | None
    fields: dict[str, tuple[str, str, bool]] = field(default_factory=dict)
    status: str = ""        # "ok" | "mismatch" | "no_excel" | "no_pdf" | "error"
    status_label: str = ""


# ------------------------------- TIỆN ÍCH CHUẨN HÓA -------------------------------

def norm_number(value: Any) -> str:
    """'00807420' / '807420' / 807420.0 -> '807420' (bỏ ký tự không phải số, bỏ số 0 ở đầu)."""
    digits = re.sub(r"[^\d]", "", str(value))
    return str(int(digits)) if digits else ""


def norm_amount(value: Any) -> str:
    """'97.000' / '97000' / 97000.0 / '2.122.200' -> '97000' / '2122200'."""
    s = str(value).strip()
    if s.endswith(".0"):          # số thực từ Excel: 97000.0
        s = s[:-2]
    return s.replace(".", "").replace(",", "")


def norm_date(value: Any) -> str:
    """Chuẩn hóa ngày về dạng DD.MM.YYYY (chấp nhận datetime hoặc chuỗi '04.08.2026')."""
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
    """Tìm tên cột chứa một trong các từ khóa (không phân biệt hoa thường)."""
    for col in df.columns:
        low = str(col).lower()
        if any(k.lower() in low for k in keywords):
            return col
    return None


def load_excel_rows(excel_path: str, sheet_name: str) -> list[ExcelRow]:
    """Đọc bảng tổng hợp, tự động dò dòng tiêu đề, trả về danh sách dòng chi phí."""
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Không tìm thấy file Excel: {excel_path}")

    raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    header_idx = next(
        (
            i for i, row in raw.iterrows()
            if HEADER_KEYWORD in " ".join(str(c) for c in row.values).lower()
        ),
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
        if pd.isna(no):
            continue
        no_str = str(no).strip()
        no_norm = norm_number(no_str)
        if not no_norm:   # bỏ qua dòng không phải số (vd: chữ ký 'Requested by/...' nằm nhầm cột)
            continue
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
    """Trích toàn bộ văn bản từ PDF (nhiều trang)."""
    parts: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                parts.append(text)
    return "\n".join(parts)


def extract_invoice_no(text: str, filename: str) -> str:
    """
    Lấy số hóa đơn, hỗ trợ cả 2 kiểu bố cục:
      - Số SAU nhãn: 'Số (Invoice No): 00006486' | 'Số: 00095512'
      - Số TRƯỚC nhãn: '807420' / 'Số (No.):' (Grab, DHL, Starbucks)
    Dự phòng cuối cùng: lấy từ tên file 'Inv_00807420'.
    """
    # 1) Số trên cùng dòng với nhãn: "Số (Invoice No): 00006486" / "Số(No):12345"
    m = re.search(r"Số\s*\((?:Invoice\s*)?No\.?\)\s*:?\s*(\d+)", text, re.IGNORECASE)
    if m:
        return m.group(1)

    # 2) Kiểu Koi: "Số: 00095512" (không có nhãn tiếng Anh)
    m = re.search(r"Số\s*:\s*(\d+)", text)
    if m:
        return m.group(1)

    # 3) Số nằm trên dòng ngay TRƯỚC nhãn: "807420\nSố (No.):"
    m = re.search(r"(\d{3,})\s*\n\s*Số\s*\(", text)
    if m:
        return m.group(1)

    # 4) Dự phòng từ tên file: "12. Inv_00873795_Grab.pdf" -> 00873795
    m = re.search(r"Inv[_\- ]?0*(\d+)", filename, re.IGNORECASE)
    return m.group(1) if m else ""


AMOUNT_PATTERNS: list[tuple[re.Pattern[str], bool]] = [
    # Grab (template cũ): "Tổng cộng số tiền đã có thuế GTGT:" + dòng kế "97.000"
    (re.compile(r"Tổng cộng số tiền đã có thuế GTGT\s*:?\s*\n?\s*([\d\.]+)", re.IGNORECASE), False),
    # Grab (template mới): "Tổng cộng hóa đơn (Invoice total): 43.000"
    (re.compile(r"Tổng cộng hóa đơn\s*(?:\([^)]*\))?\s*:?\s*([\d\.]+)", re.IGNORECASE), False),
    # Nasco / DHL / Xanh SM / Green Fast / BICOM:
    #   "Tổng cộng tiền thanh toán (Grand total): 271.269"        (Nasco/DHL, 1 số)
    #   "Tổng cộng tiền thanh toán (Grand total): 122.877 9.123 132.000"
    #     (Xanh SM có thêm cột 'Thành tiền sau thuế' -> lấy số CUỐI)
    (re.compile(r"Tổng cộng tiền thanh toán\s*(?:\([^)]*\))?\s*:?\s*([\d\.]+(?:[ \t]+[\d\.]+)*)", re.IGNORECASE), True),
    # Koi / Starbucks / Petro: "Tổng tiền thanh toán: 149.000" | "...(Total of payment): 309.000"
    (re.compile(r"Tổng tiền thanh toán\s*(?:\([^)]*\))?\s*:?\s*([\d\.]+)", re.IGNORECASE), False),
    # Takahiro / SACO: "Tổng cộng: 664.000 53.120 717.120" | "Tổng cộng(Total): 232.407 18.593 251.000"
    #   -> lấy số cuối (cột 'Cộng tiền thanh toán' / giá trị sau thuế)
    (re.compile(r"Tổng cộng\s*(?:\([^)]*\))?\s*:?\s*([\d\.]+(?:[ \t]+[\d\.]+)*)"), True),
]


def extract_amount(text: str) -> str:
    """
    Lấy tổng tiền thanh toán (đã bao gồm thuế) từ các mẫu phổ biến.

    Nếu dòng tổng cộng có nhiều số trên cùng một dòng (vd: thành tiền trước thuế,
    tiền thuế, thành tiền sau thuế), ưu tiên lấy số CUỐI — giá trị cột
    'Thành tiền sau thuế' thay vì cột 'Thành tiền' (chưa thuế).
    """
    for pattern, take_last in AMOUNT_PATTERNS:
        m = pattern.search(text)
        if m:
            numbers = re.findall(r"[\d\.]+", m.group(1))
            if numbers:
                return numbers[-1] if take_last else numbers[0]
    return ""


DATE_RE = re.compile(
    r"Ngày\s*(?:\([^)]*\))?\s*(\d{1,2})\s*tháng\s*(?:\([^)]*\))?\s*(\d{1,2})\s*năm\s*(?:\([^)]*\))?\s*(\d{4})",
    re.IGNORECASE,
)


def extract_date(text: str) -> str:
    """Ngày lập hóa đơn -> 'DD.MM.YYYY' (hỗ trợ 'Ngày (date)...' lẫn 'Ngày 21 tháng 08 năm 2026')."""
    m = DATE_RE.search(text)
    if m:
        d, mo, y = m.groups()
        return f"{int(d):02d}.{int(mo):02d}.{y}"
    return ""


def extract_company(text: str) -> str:
    """Tên đơn vị mua hàng (bỏ qua 'Tên đơn vị bán hàng')."""
    m = re.search(r"Tên đơn vị\s*\('?Company'?s?\s*name\)\s*:\s*(.+)", text, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    m = re.search(r"Tên đơn vị\s*:\s*(.+)", text)  # kiểu Koi, không có nhãn tiếng Anh
    if m:
        return m.group(1).strip()
    # Dự phòng: PDF không có nhãn rõ ràng (vd: Starbucks) — dò tên công ty mua hàng
    if "ACCLIME" in text.upper():
        return EXPECTED_COMPANY
    return ""


def extract_taxcode(text: str) -> str:
    """
    Lấy mã số thuế NGƯỜI MUA (ACCLIME). PDF thường chứa 2 mã số thuế
    (người bán + người mua) nên ưu tiên mã khớp EXPECTED_TAXCODE,
    nếu không có thì lấy mã xuất hiện cuối cùng.
    """
    codes = [
        re.sub(r"[^\d]", "", m.group(1))
        for m in re.finditer(r"Mã số thuế\s*(?:\([^)]*\))?\s*:\s*([\d\s\-]+)", text)
    ]
    codes = [c for c in codes if c]
    if not codes:
        # Dự phòng: PDF tách nhãn và giá trị ra dòng khác nhau (Starbucks) —
        # dò các số 10 chữ số (mã số thuế) và ưu tiên mã của người mua.
        codes = re.findall(r"\b\d{10}\b", text)
    if not codes:
        return ""
    return next((c for c in codes if c == EXPECTED_TAXCODE), codes[-1])


SIGNATURE_RE = re.compile(r"(Signature Valid|Ký bởi|ký điện tử|Signed by|đã ký)", re.IGNORECASE)


def extract_signature(text: str) -> bool:
    """Kiểm tra sự hiện diện của khối chữ ký / chữ ký số trong hóa đơn."""
    return bool(SIGNATURE_RE.search(text))


def parse_pdf(pdf_path: str) -> PdfData:
    """Trích xuất toàn bộ thông tin cần thiết từ một file PDF."""
    data = PdfData(path=pdf_path)
    try:
        text = extract_pdf_text(pdf_path)
    except Exception as exc:  # noqa: BLE001 - lỗi file riêng lẻ không làm dừng toàn bộ
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
    """Tìm dòng Excel khớp số chứng từ (so sánh sau khi bỏ số 0 ở đầu)."""
    if not pdf_no_norm:
        return None
    return next((r for r in excel_rows if r.no_norm == pdf_no_norm), None)


def _fold_diacritics(s: str) -> str:
    """Bỏ dấu tiếng Việt để so sánh không phân biệt 'CÔNG' vs 'CONG'."""
    return "".join(
        c for c in unicodedata.normalize("NFD", s.lower())
        if unicodedata.category(c) != "Mn"
    )


def compare(excel: ExcelRow, pdf: PdfData) -> dict[str, tuple[str, str, bool]]:
    """So sánh từng tiêu chí, trả về dict: tiêu chí -> (giá trị Excel, giá trị PDF, khớp?)."""
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
    """Tạo kết quả đối chiếu cho một PDF."""
    if pdf.error:
        return Result(pdf=pdf, excel=excel, status="error", status_label="Lỗi đọc file")
    if excel is None:
        return Result(pdf=pdf, excel=None, status="no_excel", status_label="Không có trong Excel")
    fields = compare(excel, pdf)
    all_ok = all(ok for _, _, ok in fields.values())
    status = "ok" if all_ok else "mismatch"
    label = "✅ Khớp" if all_ok else "❌ Lệch"
    return Result(pdf=pdf, excel=excel, fields=fields, status=status, status_label=label)


# ------------------------------- XUẤT KẾT QUẢ -------------------------------

def export_csv(results: list[Result], excel_missing: list[ExcelRow]) -> str:
    """Tạo nội dung CSV (kèm BOM để Excel hiển thị đúng tiếng Việt)."""
    import csv
    header = ["STT", "File PDF", "Số CT (Excel)", "Số CT (PDF)", "Ngày (Excel)", "Ngày (PDF)",
              "Thành tiền (Excel)", "Thành tiền (PDF)", "Tên công ty", "Mã số thuế", "Chữ ký", "Kết luận", "Ghi chú"]
    rows: list[list[str]] = []
    for r in results:
        if r.pdf is None:
            continue
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
    return "\ufeff" + buf.getvalue()  # BOM để Excel nhận diện UTF-8


def render_console(results: list[Result], excel_missing: list[ExcelRow]) -> None:
    """In bảng tổng hợp ra console."""
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
        if r.pdf is None:
            continue
        stt = r.excel.stt if r.excel else "—"
        no_ex = r.excel.no if r.excel else "—"
        no_pdf = r.pdf.no
        date = r.pdf.date
        amount = r.pdf.amount or "—"
        tax = r.pdf.tax or "—"
        sign = "✅" if r.pdf.has_signature else "❌"
        table.append([stt, os.path.basename(r.pdf.path), no_ex, no_pdf, date, amount, tax, sign, r.status_label])
    print(tabulate(table, headers=headers, tablefmt="fancy_grid"))
    print(f"\n📝 Chi tiết từng trường (rà soát tay): mở file {OUTPUT_HTML} bằng trình duyệt.\n")


def render_html(results: list[Result], excel_missing: list[ExcelRow]) -> str:
    """Tạo báo cáo HTML (bảng tìm kiếm, lọc trạng thái, sắp xếp theo cột)."""
    rows_json: list[dict[str, Any]] = []

    for r in results:
        if r.pdf is None:
            continue
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
        "total": len(results),
        "ok": sum(1 for r in results if r.status == "ok"),
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
<title>Đối chiếu hóa đơn PDF ↔ Excel — {SHEET_NAME}</title>
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
  <h1>📊 Đối chiếu hóa đơn PDF ↔ Excel — {html.escape(SHEET_NAME)}</h1>
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
    """Tìm tất cả file PDF (đệ quy) trong thư mục."""
    pdfs: list[str] = []
    for root, _dirs, files in os.walk(folder):
        for f in sorted(files):
            if f.lower().endswith(".pdf") and not f.startswith("~$"):
                pdfs.append(os.path.join(root, f))
    return sorted(pdfs)


def parse_args() -> argparse.Namespace:
    """Đọc tham số dòng lệnh, cho phép dùng lại cho các kỳ claim khác."""
    parser = argparse.ArgumentParser(description="Đối chiếu hóa đơn PDF với bảng tổng hợp chi phí Excel.")
    parser.add_argument("--excel", default=EXCEL_FILE, help=f"Đường dẫn file Excel (mặc định: {EXCEL_FILE})")
    parser.add_argument("--sheet", default=SHEET_NAME, help=f"Tên sheet chứa bảng tổng hợp (mặc định: {SHEET_NAME})")
    parser.add_argument("--pdf-folder", default=PDF_FOLDER, help=f"Thư mục chứa PDF (mặc định: {PDF_FOLDER})")
    parser.add_argument("--output", default=OUTPUT_HTML, help=f"File báo cáo HTML (mặc định: {OUTPUT_HTML})")
    parser.add_argument("--csv", default=None, help="File CSV xuất kết quả (mặc định: không xuất)")
    parser.add_argument("--no-open", action="store_true", help="Không tự động mở báo cáo HTML trong trình duyệt")
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    excel_rows = load_excel_rows(args.excel, args.sheet)
    print(f"📋 Đã đọc {len(excel_rows)} dòng chi phí từ sheet '{args.sheet}' của {args.excel}")

    pdf_files = find_pdfs(args.pdf_folder)
    print(f"📄 Tìm thấy {len(pdf_files)} file PDF trong {args.pdf_folder}\n")

    results: list[Result] = []
    matched_nos: set[str] = set()

    for pdf_path in pdf_files:
        pdf = parse_pdf(pdf_path)
        excel = match_excel_row(excel_rows, pdf.no_norm) if not pdf.error else None
        if excel is not None:
            matched_nos.add(excel.no_norm)
        results.append(build_result(pdf, excel))

    excel_missing = [r for r in excel_rows if r.no_norm not in matched_nos]

    render_console(results, excel_missing)

    with open(args.output, "w", encoding="utf-8") as f:
        f.write(render_html(results, excel_missing))
    print(f"💾 Đã xuất báo cáo HTML: {os.path.abspath(args.output)}")

    if args.csv:
        with open(args.csv, "w", encoding="utf-8", newline="") as f:
            f.write(export_csv(results, excel_missing))
        print(f"📄 Đã xuất CSV: {os.path.abspath(args.csv)}")

    if not args.no_open and sys.platform.startswith("win"):
        try:
            webbrowser.open(os.path.abspath(args.output).replace("\\", "/"))
        except Exception:  # noqa: BLE001 - không chặn luồng chính nếu không mở được
            pass


if __name__ == "__main__":
    main()
