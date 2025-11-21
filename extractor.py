# Updated extractor.py with generalized parsing, raw dump handling, and flexible logic
import re
import argparse
from pathlib import Path
import pandas as pd
import pdfplumber
import logging
import datetime

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

def extract_text_from_pdf(pdf_path):
    logging.info(f"Extracting text from: {pdf_path}")
    all_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                all_text.append(txt)
    return "\n".join(all_text)

# NEW: Generic key-value detector
def detect_key_value_lines(text):
    rows = []
    lines = text.split("\n")
    kv_pattern = re.compile(r"^([^:]+):\s*(.*)$")

    for line in lines:
        m = kv_pattern.match(line.strip())
        if m:
            key = m.group(1).strip()
            value = m.group(2).strip()
            rows.append({"Key": key, "Value": value, "Comments": "Detected from generic key:value structure"})
    return rows

# NEW: Flexible extraction of dates, numbers, names
def detect_patterns(text):
    rows = []

    # Date detection
    date_patterns = [
        r"\b\d{4}-\d{2}-\d{2}\b",
        r"\b\d{1,2} [A-Za-z]+ \d{4}\b",
        r"[A-Za-z]+ \d{1,2}, \d{4}"
    ]

    comment = "Identified automatically from detected date patterns"

    for p in date_patterns:
        matches = re.findall(p, text)
        for m in matches:
            rows.append({"Key": "Detected Date", "Value": m, "Comments": comment})

    # Number detection (years, amounts)
    num_matches = re.findall(r"\b\d{4}\b", text)
    for n in num_matches:
        rows.append({"Key": "Detected Number", "Value": n, "Comments": "Auto-detected numeric entity"})

    return rows

# NEW: Raw dump fallback
def raw_dump_section(text):
    return [{"Key": "RAW_TEXT", "Value": text, "Comments": "Unprocessed content dump (100% capture)"}]

# Master builder
def build_rows_from_text(text):
    rows = []
    rows.extend(detect_key_value_lines(text))
    rows.extend(detect_patterns(text))
    rows.extend(raw_dump_section(text))
    return rows


def run(pdf_path, output_xlsx):
    text = extract_text_from_pdf(pdf_path)

    rows = build_rows_from_text(text)
    df = pd.DataFrame(rows)[["Key", "Value", "Comments"]]

    out_path = Path(output_xlsx)
    df.to_excel(out_path, index=False)
    logging.info(f"Saved structured Excel to {out_path.resolve()}")
    return df

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", "-p", type=str, required=True, help="Path to input PDF")
    parser.add_argument("--out", "-o", type=str, default="Output.xlsx", help="Output Excel path")
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        logging.error(f"PDF not found: {pdf_path}")
        raise SystemExit(1)

    df = run(pdf_path, args.out)
    print("\nPreview of output:\n")
    print(df.to_string(index=False))
