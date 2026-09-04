"""
app.py — Website đối chiếu hóa đơn PDF với bảng tổng hợp chi phí Excel.

Upload file Excel + file PDF (hoặc file ZIP chứa PDF) → bấm nút → xem kết quả
dạng bảng trên web, tải báo cáo HTML/CSV về. Không cần chạy lệnh.

Cách chạy local:
    pip install streamlit pandas pdfplumber tabulate openpyxl
    streamlit run app.py
"""

from __future__ import annotations

import io
import os
import tempfile
import zipfile

import pandas as pd
import streamlit as st

import invoice_checker as ic

st.set_page_config(page_title="Đối chiếu hóa đơn PDF ↔ Excel", page_icon="📊", layout="wide")

# ---------------- Tiêu đề ----------------
st.title("📊 Đối chiếu hóa đơn PDF ↔ Excel")
st.caption("Upload file Excel (bảng tổng hợp chi phí) và các file PDF hóa đơn, bấm **Chạy đối chiếu** để kiểm tra tự động.")

# ---------------- Cấu hình nâng cao (mặc định khớp với claim hiện tại) ----------------
with st.expander("⚙️ Cấu hình nâng cao"):
    c1, c2 = st.columns(2)
    expected_company = c1.text_input("Tên công ty người mua (so sánh)", value=ic.EXPECTED_COMPANY)
    expected_taxcode = c2.text_input("Mã số thuế người mua (so sánh)", value=ic.EXPECTED_TAXCODE)


# ---------------- Bước 1: upload Excel ----------------
st.subheader("1️⃣ Tải lên file Excel (bảng tổng hợp)")
excel_file = st.file_uploader("File Excel (.xlsx)", type=["xlsx", "xls"], key="excel")

sheet_name = "Claim"
if excel_file is not None:
    try:
        sheets = pd.ExcelFile(excel_file).sheet_names
        sheet_name = st.selectbox("Chọn sheet chứa bảng chi phí", sheets, index=sheets.index("Claim") if "Claim" in sheets else 0)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Không đọc được file Excel: {exc}")

# ---------------- Bước 2: upload PDF ----------------
st.subheader("2️⃣ Tải lên hóa đơn PDF")
uploaded_pdfs = st.file_uploader(
    "Chọn nhiều file PDF, hoặc 1 file ZIP chứa PDF",
    type=["pdf", "zip"],
    accept_multiple_files=True,
    key="pdfs",
)

# ---------------- Bước 3: chạy đối chiếu ----------------
run_clicked = st.button("🚀 Chạy đối chiếu", type="primary", disabled=excel_file is None or not uploaded_pdfs)

if run_clicked:
    if excel_file is None or not uploaded_pdfs:
        st.warning("Vui lòng upload đủ file Excel và file PDF.")
        st.stop()

    # Lưu file tạm để dùng lại đúng logic đã kiểm thử (pdfplumber, pandas)
    with tempfile.TemporaryDirectory() as tmpdir:
        excel_path = os.path.join(tmpdir, "claim.xlsx")
        with open(excel_path, "wb") as f:
            f.write(excel_file.getbuffer())

        pdf_paths: list[str] = []
        for up in uploaded_pdfs:
            name = up.name
            if name.lower().endswith(".zip"):
                with zipfile.ZipFile(io.BytesIO(up.getbuffer())) as z:
                    for member in z.namelist():
                        if member.lower().endswith(".pdf"):
                            target = os.path.join(tmpdir, os.path.basename(member))
                            with z.open(member) as src, open(target, "wb") as dst:
                                dst.write(src.read())
                            pdf_paths.append(target)
            else:
                target = os.path.join(tmpdir, name)
                with open(target, "wb") as f:
                    f.write(up.getbuffer())
                pdf_paths.append(target)

        # Ghi đè cấu hình so sánh từ giao diện
        ic.EXPECTED_COMPANY = expected_company.strip()
        ic.EXPECTED_TAXCODE = expected_taxcode.strip()

        # Đọc Excel
        try:
            excel_rows = ic.load_excel_rows(excel_path, sheet_name)
        except Exception as exc:  # noqa: BLE001
            st.error(f"Lỗi đọc Excel: {exc}")
            st.stop()

        # Xử lý từng PDF
        results: list[ic.Result] = []
        matched_nos: set[str] = set()
        progress = st.progress(0.0)
        status_box = st.empty()

        for i, pdf_path in enumerate(pdf_paths):
            status_box.info(f"Đang xử lý: {os.path.basename(pdf_path)} ({i + 1}/{len(pdf_paths)})")
            pdf = ic.parse_pdf(pdf_path)
            excel = ic.match_excel_row(excel_rows, pdf.no_norm) if not pdf.error else None
            if excel is not None:
                matched_nos.add(excel.no_norm)
            results.append(ic.build_result(pdf, excel))
            progress.progress((i + 1) / len(pdf_paths))

        status_box.empty()
        excel_missing = [r for r in excel_rows if r.no_norm not in matched_nos]

        # ---------------- Kết quả ----------------
        st.subheader("📋 Kết quả đối chiếu")
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Tổng PDF", len(results))
        m2.metric("✅ Khớp", sum(1 for r in results if r.status == "ok"))
        m3.metric("❌ Lệch", sum(1 for r in results if r.status == "mismatch"))
        m4.metric("⚠️ Không có trong Excel", sum(1 for r in results if r.status == "no_excel"))
        m5.metric("📌 Excel thiếu PDF", len(excel_missing))

        # Bảng tương tác (tìm kiếm, lọc, sắp xếp) — dùng lại báo cáo HTML
        html_report = ic.render_html(results, excel_missing)
        st.components.v1.html(html_report, height=760, scrolling=True)

        # ---------------- Tải về ----------------
        st.subheader("⬇️ Tải kết quả về")
        d1, d2 = st.columns(2)
        d1.download_button(
            "📄 Tải báo cáo HTML", html_report,
            file_name="invoice_check_report.html", mime="text/html",
        )
        csv_data = ic.export_csv(results, excel_missing)
        d2.download_button(
            "📊 Tải CSV (mở bằng Excel)", csv_data,
            file_name="invoice_check_report.csv", mime="text/csv",
        )

        # Cảnh báo các trường hợp cần rà soát tay
        problems = [r for r in results if r.status in ("mismatch", "no_excel", "error")] + \
                   [ic.Result(pdf=None, excel=r, status="no_pdf", status_label="Thiếu file PDF") for r in excel_missing]
        if problems:
            st.warning(f"Có {len(problems)} trường hợp cần kiểm tra tay — xem chi tiết trong bảng phía trên (hàng màu đỏ/vàng).")

st.markdown("---")
st.caption("Công cụ nội bộ: dữ liệu được xử lý ngay trong phiên, không lưu trên server sau khi chạy xong.")