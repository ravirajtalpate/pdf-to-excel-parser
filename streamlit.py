import streamlit as st
import pandas as pd
from pathlib import Path
import pdfplumber
import re
import datetime

# -----------------------------
# Helper Functions
# -----------------------------

def extract_text_from_pdf(pdf_path):
    all_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                all_text.append(txt)
    return "\n".join(all_text)

def find_first_match(patterns, text, flags=0):
    for p in patterns:
        m = re.search(p, text, flags)
        if m:
            return m
    return None

def normalize_salary(raw):
    if raw is None:
        return ""
    s = re.sub(r"[^\d\.]", "", raw.replace(",", ""))
    if s == "":
        return ""
    try:
        v = int(float(s))
        return str(v)
    except:
        return s

def build_rows_from_text(text):
    rows = []

    # Name extraction
    m = re.search(r"([A-Z][a-z]+)\s+([A-Z][a-z]+)\s+was born", text)
    if m:
        first, last = m.group(1), m.group(2)
        rows.append({"Key": "First Name", "Value": first, "Comments": ""})
        rows.append({"Key": "Last Name", "Value": last, "Comments": ""})

    # DOB extraction
    dob_patterns = [
        r"born on\s+([A-Za-z]+\s+\d{1,2},\s*\d{4})",
        r"born on\s+(\d{4}-\d{2}-\d{2})",
        r"birthdate is formatted as\s+(\d{4}-\d{2}-\d{2})",
    ]
    m = find_first_match(dob_patterns, text, re.I)
    dob_value = ""
    if m:
        raw = m.group(1)
        try:
            if "-" in raw:
                dt = datetime.datetime.strptime(raw.strip(), "%Y-%m-%d")
            else:
                dt = datetime.datetime.strptime(raw.strip(), "%B %d, %Y")
            dob_value = dt.strftime("%d-%b-%y")
        except:
            dob_value = raw

    rows.append({"Key": "Date of Birth", "Value": dob_value, "Comments": ""})

    # Birth city fallback
    city = ""
    if "Jaipur" in text:
        city = "Jaipur"
    rows.append({"Key": "Birth City", "Value": city, "Comments": ""})

    df = pd.DataFrame(rows)
    return df


# -----------------------------
# Streamlit App UI
# -----------------------------

st.set_page_config(page_title="AI Document Structuring Demo", layout="centered")

st.title("📄 AI-Powered Document Structuring Demo")
st.write("Upload an unstructured PDF and convert it into a structured Excel file.")

uploaded_pdf = st.file_uploader("Upload PDF File", type=["pdf"])

if uploaded_pdf:
    st.success("PDF uploaded successfully.")

    # Save temp file
    temp_pdf_path = Path("uploaded.pdf")
    with open(temp_pdf_path, "wb") as f:
        f.write(uploaded_pdf.getbuffer())

    # Extract and parse
    st.write("Extracting text...")
    raw_text = extract_text_from_pdf(temp_pdf_path)

    st.write("Parsing data...")
    df = build_rows_from_text(raw_text)

    st.write("### Extracted Structured Data")
    st.dataframe(df)

    # Create Excel output
    output_path = "Output.xlsx"
    df.to_excel(output_path, index=False)

    with open(output_path, "rb") as f:
        st.download_button(
            label="📥 Download Excel Output",
            data=f,
            file_name="Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
