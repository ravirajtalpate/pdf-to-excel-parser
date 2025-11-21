import pdfplumber
import pandas as pd
import re

# ---------------------------------------------------------
# STEP 1: EXTRACT ALL RAW TEXT FROM PDF
# ---------------------------------------------------------

def extract_pdf_text(pdf_path):
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            full_text += page.extract_text() + "\n"
    return full_text


# ---------------------------------------------------------
# STEP 2: IDENTIFY KEY:VALUE PAIRS
# This is generic logic since your actual data PDF is missing.
# Adjust regex when you upload the real PDF.
# ---------------------------------------------------------

def extract_key_value_pairs(text):
    pairs = []
    comments = []

    lines = [l.strip() for l in text.split("\n") if l.strip()]

    for line in lines:
        # Try simple key:value detection
        if ":" in line:
            key, value = line.split(":", 1)
            pairs.append((key.strip(), value.strip()))
            comments.append("")  # no comment for direct match
        else:
            # Non key:value lines → put as comment-only rows
            pairs.append(("", ""))
            comments.append(line)

    return pairs, comments


# ---------------------------------------------------------
# STEP 3: SAVE INTO OUTPUT EXCEL
# ---------------------------------------------------------

def save_to_excel(pairs, comments, output_path):
    df = pd.DataFrame({
        "Key": [p[0] for p in pairs],
        "Value": [p[1] for p in pairs],
        "Comments": comments
    })

    df.to_excel(output_path, index=False)
    print("Saved:", output_path)


# ---------------------------------------------------------
# MAIN PIPELINE
# ---------------------------------------------------------

if __name__ == "__main__":
    pdf_path = "Data Input.pdf"       # change when you upload it
    output_path = "Output.xlsx"

    print("Extracting text...")
    text = extract_pdf_text(pdf_path)

    print("Extracting key:value pairs...")
    pairs, comments = extract_key_value_pairs(text)

    print("Saving output...")
    save_to_excel(pairs, comments, output_path)

    print("Done.")
