"""
Streamlit Web Application for PDF to Excel Extraction (No API Version)
"""

import streamlit as st
import PyPDF2
import pandas as pd
import re
import io
from datetime import datetime
from typing import List, Dict
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Page configuration
st.set_page_config(
    page_title="PDF to Excel Extractor",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        text-align: center;
        margin-bottom: 2rem;
    }
    </style>
""", unsafe_allow_html=True)


# --------------------------------------------------------------
# BASIC LOGIC FOR NON-AI EXTRACTION
# --------------------------------------------------------------

def extract_text_from_pdf(pdf_file):
    """Extract raw text from PDF."""
    text = ""
    reader = PyPDF2.PdfReader(pdf_file)
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text


def extract_key_value_pairs(text: str) -> List[Dict]:
    """
    A simple heuristic:
    • Lines with ":" become Key/Value
    • Everything else becomes a comment row
    """

    structured = []
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    for line in lines:
        if ":" in line:
            key, value = line.split(":", 1)
            structured.append({
                "Key": key.strip(),
                "Value": value.strip(),
                "Comments": ""
            })
        else:
            structured.append({
                "Key": "",
                "Value": "",
                "Comments": line
            })

    return structured


def create_excel_output(structured_data: List[Dict]) -> bytes:
    """Create final Excel with formatting."""

    df = pd.DataFrame(structured_data)
    df.insert(0, '#', range(1, 1 + len(df)))

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Output', startrow=1)

        workbook = writer.book
        sheet = writer.sheets['Output']

        # Styles
        header_fill = PatternFill(start_color="D3D3D3", fill_type="solid")
        header_font = Font(bold=True, size=11)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Style header
        for cell in sheet[2]:
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal="center")

        # Style data cells
        for row in sheet.iter_rows(min_row=3, max_row=sheet.max_row, min_col=1, max_col=4):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical="top")

        # Column widths
        sheet.column_dimensions['A'].width = 5
        sheet.column_dimensions['B'].width = 30
        sheet.column_dimensions['C'].width = 40
        sheet.column_dimensions['D'].width = 80

    output.seek(0)
    return output.getvalue()


# --------------------------------------------------------------
# STREAMLIT UI
# --------------------------------------------------------------

def main():

    st.markdown('<div class="main-header">📄 PDF to Excel Extractor (No AI)</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Simple key:value extraction without using any API</div>', unsafe_allow_html=True)

    st.header("📤 Upload PDF")
    pdf = st.file_uploader("Select a PDF file", type=["pdf"])

    if pdf:

        st.subheader("📄 Extracted Text Preview")

        text = extract_text_from_pdf(pdf)
        st.text(text[:1200] + ("..." if len(text) > 1200 else ""))

        if st.button("🚀 Convert to Excel", type="primary"):
            structured = extract_key_value_pairs(text)
            excel_bytes = create_excel_output(structured)

            df = pd.DataFrame(structured)
            df.insert(0, '#', range(1, 1 + len(df)))

            st.subheader("📊 Preview")
            st.dataframe(df, use_container_width=True, height=400)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

            st.download_button(
                label="📥 Download Excel",
                data=excel_bytes,
                file_name=f"output_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


if __name__ == "__main__":
    main()
