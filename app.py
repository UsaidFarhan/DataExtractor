import subprocess
import sys
import os

req_file = os.path.join(os.path.dirname(__file__), "requirements.txt")
if os.path.exists(req_file):
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", req_file, "-q"])

import re
import io
import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import streamlit as st

PATTERNS = {
    "Employee Code": [
        r"Employee\s*Number[:\s]+(\d+)",
        r"Employee\s*Code[:\s]+(\d+)",
        r"Emp\.?\s*No\.?[:\s]+(\d+)",
        r"Ref:\s*\(per\)\s*/\s*file\s+(\d+)",
    ],
    "New Gross Salary": [
        r"(?:gross\s*salary\s*from\s*Rs\.?\s*[₨]?\s*[\d,\.]+\s*[\/\-]*\s*to\s*Rs\.?\s*[₨]?\s*)([\d,\.]+)",
        r"(?:gross\s*salary\s*in\s*your\s*new\s*grade\s*will\s*be\s*Rs\.?)\s*[₨]?\s*([\d,\.]+)",
        r"(?:revised\s*gross\s*salary\s*will\s*be\s*Rs\.?)\s*[₨]?\s*([\d,\.]+)",
        r"(?:monthly\s*gross\s*salary\s*will\s*be\s*Rs\.?)\s*[₨]?\s*([\d,\.]+)",
        r"(?:new\s*gross\s*salary\s*will\s*be\s*Rs\.?)\s*[₨]?\s*([\d,\.]+)",
        r"(?:gross\s*salary\s*will\s*be\s*Rs\.?)\s*[₨]?\s*([\d,\.]+)",
        r"(?:salary\s*will\s*be\s*Rs\.?)\s*[₨]?\s*([\d,\.]+)",
        r"(?:will\s*be\s*Rs\.?)\s*[₨]?\s*([\d,\.]+)",
        r"(?:to\s*Rs\.?\s*[₨]?\s*)([\d,\.]+)\s*[\/\-]*\s*per\s*month",
    ],
    "Difference": [
        r"(?:gross\s*)?increase\s*of\s*Rs\.?\s*[₨]?\s*(-?[\d,\.]+)\s*[\/\-]*",
        r"increment\s*of\s*Rs\.?\s*[₨]?\s*(-?[\d,\.]+)\s*[\/\-]*",
        r"raise\s*of\s*Rs\.?\s*[₨]?\s*(-?[\d,\.]+)\s*[\/\-]*",
        r"salary\s*increase\s*(?:of|:)\s*Rs\.?\s*[₨]?\s*(-?[\d,\.]+)\s*[\/\-]*",
        r"revision\s*of\s*Rs\.?\s*[₨]?\s*(-?[\d,\.]+)\s*[\/\-]*",
        r"increment\s*(?:amount\s*)?(?:is|:)\s*Rs\.?\s*[₨]?\s*(-?[\d,\.]+)\s*[\/\-]*",
    ],
    "Date Effective From": [
        r"effective\s*from\s+([A-Za-z]+\s+\d{1,2},\s*\d{4})",
        r"effective\s*from\s+(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})",
        r"with\s*effect\s*from\s+([A-Za-z]+\s+\d{1,2},?\s*\d{4})",
        r"with\s*effect\s*from\s+(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})",
        r"effective\s*date[:\s]+([A-Za-z]+\s+\d{1,2},\s*\d{4})",
        r"effective\s*date[:\s]+(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})",
        r"w\.?e\.?f\.?\s*([A-Za-z]+\s+\d{1,2},\s*\d{4})",
        r"w\.?e\.?f\.?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})",
    ],
}

def extract_fields(text):
    results = {}
    for field, patterns in PATTERNS.items():
        matched = "NOT FOUND"
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                matched = match.group(1).strip()
                break
        results[field] = matched
    return results

def process_pdf(file_bytes, filename):
    results = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        total = len(pdf.pages)
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            fields = extract_fields(text)
            label = filename if total == 1 else f"{filename} — Page {i}"
            results.append((label, fields))
    return results

def build_excel(data_rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Salary Data"

    headers = ["Source", "Employee Code", "New Gross Salary", "Difference", "Date Effective From"]

    header_fill = PatternFill("solid", start_color="2F5496", end_color="2F5496")
    header_font = Font(bold=True, color="FFFFFF", name="Arial")
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    row_font = Font(name="Arial", size=10)
    for row_idx, row in enumerate(data_rows, start=2):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = row_font
            if row_idx % 2 == 0:
                cell.fill = PatternFill("solid", start_color="DCE6F1", end_color="DCE6F1")

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ── STREAMLIT UI ────────────────────────────────────────────────
st.set_page_config(page_title="Salary Letter Extractor", page_icon="📄", layout="centered")

st.title("📄 Salary Letter Extractor")
st.markdown("Upload one or more PDF files and click **Extract** to generate an Excel file.")

uploaded_files = st.file_uploader(
    "Upload PDF file(s)",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("Extract & Generate Excel", type="primary"):
        all_rows = []
        total_pages = 0
        not_found_count = 0

        with st.spinner("Processing..."):
            for uploaded_file in uploaded_files:
                file_bytes = uploaded_file.read()
                page_results = process_pdf(file_bytes, uploaded_file.name)
                total_pages += len(page_results)

                for label, fields in page_results:
                    row = [
                        label,
                        fields["Employee Code"],
                        fields["New Gross Salary"],
                        fields["Difference"],
                        fields["Date Effective From"],
                    ]
                    all_rows.append(row)
                    not_found_count += sum(1 for v in fields.values() if v == "NOT FOUND")

        if not all_rows:
            st.error("No data could be extracted from the uploaded files.")
        else:
            fully_extracted = sum(1 for row in all_rows if "NOT FOUND" not in row)

            col1, col2, col3 = st.columns(3)
            col1.metric("Letters Processed", total_pages)
            col2.metric("Fully Extracted", fully_extracted)
            col3.metric("Fields Not Found", not_found_count)

            excel_buffer = build_excel(all_rows)
            base_name = os.path.splitext(uploaded_files[0].name)[0]
            output_name = f"SalaryData-{base_name}.xlsx"

            st.download_button(
                label="⬇️ Download Excel",
                data=excel_buffer,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
