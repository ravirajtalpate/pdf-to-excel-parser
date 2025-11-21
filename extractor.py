"""
pdf_to_excel.py

Simple, dependency-light PDF -> Excel extractor following the assignment rules:
- Extracts raw text from PDF
- Uses heuristic rules to detect key:value pairs (preserves original wording)
- Collects additional contextual lines into a "Comments" field
- Writes a well-formatted Excel file

Usage (CLI):
    python pdf_to_excel.py input.pdf output.xlsx

Or import functions in another script / Streamlit app.
"""
import sys
import re
import os
from typing import List, Dict
import json

import PyPDF2
import pandas as pd
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl import load_workbook

# --- Text extraction -------------------------------------------------------
def extract_text_from_pdf(path: str) -> str:
    """Extract full text from PDF. Returns one big string with newline separators."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"Input file not found: {path}")
    text_parts = []
    with open(path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            page_text = page.extract_text() or ""
            text_parts.append(page_text)
    return "\n".join(text_parts)


# --- Heuristic key:value parsing ------------------------------------------
KEY_COLON_RE = re.compile(r'^\s*([A-Za-z0-9][A-Za-z0-9 .&()/\-]{0,80}?)\s*:\s*(.+)$')
# fallback: lines that look like "Name — value" or "Name - value"
KEY_DASH_RE = re.compile(r'^\s*([A-Za-z0-9][A-Za-z0-9 .&()/\-]{0,80}?)\s+[—\-–]\s+(.+)$')

def detect_key_line(line: str):
    """Return (key, value) if the line looks like a key:value line, else None."""
    m = KEY_COLON_RE.match(line)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    m2 = KEY_DASH_RE.match(line)
    if m2:
        return m2.group(1).strip(), m2.group(2).strip()
    return None

def parse_text_to_pairs(raw_text: str) -> List[Dict]:
    """
    Convert raw_text to a list of {Key, Value, Comments}.
    Heuristics:
      - Lines with colon or dash separate keys and values.
      - Lines without keys are appended as comments to the current item.
      - If a long block appears with no obvious keys, we create a single 'Full text' record.
    """
    lines = [ln.rstrip() for ln in raw_text.splitlines()]
    items = []
    current = None

    for i, ln in enumerate(lines):
        if not ln.strip():
            # empty line -> treat as paragraph separator: keep but as newline in comments
            if current is not None:
                current['Comments'] += "\n"
            continue

        kv = detect_key_line(ln)
        if kv:
            # start new item
            key, val = kv
            # finalize previous
            if current:
                items.append(current)
            current = {"Key": key, "Value": val, "Comments": ""}
            continue

        # If line looks like an "All caps header" and next line is descriptive, treat as key
        if ln.strip().isupper() and i + 1 < len(lines) and lines[i+1].strip():
            # use the next non-empty line as value if it doesn't look like a key
            nxt = lines[i+1].strip()
            if not detect_key_line(nxt):
                if current:
                    items.append(current)
                current = {"Key": ln.strip().title(), "Value": nxt, "Comments": ""}
                # skip the next line (we already used it)
                continue

        # If no key detected, append line to comments of current item if exists
        if current is not None:
            if current['Comments']:
                current['Comments'] += "\n" + ln
            else:
                current['Comments'] = ln
        else:
            # no current item: create a fallback "Unstructured" item
            current = {"Key": "Unstructured", "Value": "", "Comments": ln}

    # finalize last
    if current:
        items.append(current)

    # Post-process: ensure Value is not empty where possible by pulling small first sentences into Value if Key looked like a heading
    for item in items:
        if not item['Value']:
            # try to split comments into a first-line value + rest comments if sensible
            if item['Comments']:
                split_lines = item['Comments'].splitlines()
                if len(split_lines) >= 1 and len(split_lines[0]) < 120:
                    candidate = split_lines[0].strip()
                    # move candidate to Value
                    item['Value'] = candidate
                    rest = "\n".join(split_lines[1:]).strip()
                    item['Comments'] = rest

    # Number duplicate keys like "Certifications 1, 2"
    key_count = {}
    for it in items:
        k = it['Key']
        key_count.setdefault(k, 0)
        key_count[k] += 1
        if key_count[k] > 1:
            it['Key'] = f"{k} {key_count[k]}"

    return items


# --- Excel output ----------------------------------------------------------
def create_excel_output(structured_data: List[Dict], output_path: str):
    """
    Create an Excel file with columns: #, Key, Value, Comments
    Styles header and wraps text for comments.
    """
    df = pd.DataFrame(structured_data)
    # ensure columns presence
    for col in ("Key", "Value", "Comments"):
        if col not in df.columns:
            df[col] = ""

    df = df[["Key", "Value", "Comments"]]
    df.insert(0, "#", range(1, 1 + len(df)))

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Output", index=False, startrow=1)
        workbook = writer.book
        worksheet = writer.sheets["Output"]

        # header row styling (row 2 because we started at startrow=1)
        header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
        header_font = Font(bold=True, size=11)
        border = Border(left=Side(style="thin"), right=Side(style="thin"),
                        top=Side(style="thin"), bottom=Side(style="thin"))

        for cell in worksheet[2]:
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # apply borders and wrap to data rows
        for row in worksheet.iter_rows(min_row=3, max_row=worksheet.max_row, min_col=1, max_col=4):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical="top", wrap_text=True)

        # reasonable widths
        worksheet.column_dimensions["A"].width = 5
        worksheet.column_dimensions["B"].width = 30
        worksheet.column_dimensions["C"].width = 40
        worksheet.column_dimensions["D"].width = 80

    print(f"Excel written: {output_path}")


# --- CLI runner ------------------------------------------------------------
def process_pdf_to_excel(input_pdf: str, output_xlsx: str):
    print(f"Extracting text from: {input_pdf}")
    raw = extract_text_from_pdf(input_pdf)
    print(f"Extracted {len(raw)} characters.")
    structured = parse_text_to_pairs(raw)
    print(f"Detected {len(structured)} records.")
    create_excel_output(structured, output_xlsx)
    # Also save JSON copy
    json_path = os.path.splitext(output_xlsx)[0] + ".json"
    with open(json_path, "w", encoding="utf-8") as jf:
        json.dump(structured, jf, indent=2, ensure_ascii=False)
    print(f"JSON written: {json_path}")
    return structured

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python pdf_to_excel.py input.pdf output.xlsx")
        sys.exit(1)
    inp = sys.argv[1]
    outp = sys.argv[2]
    process_pdf_to_excel(inp, outp)
