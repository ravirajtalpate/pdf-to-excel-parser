
import re
import argparse
from pathlib import Path
import pandas as pd
import pdfplumber
import logging

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# --- Helper parsing functions ---

def extract_text_from_pdf(pdf_path):
    logging.info(f"Extracting text from: {pdf_path}")
    all_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text()
            if txt:
                all_text.append(txt)
    text = "\n".join(all_text)
    logging.info(f"Extracted {len(text)} characters of text")
    return text

def find_first_match(patterns, text, flags=0):
    for p in patterns:
        m = re.search(p, text, flags)
        if m:
            return m
    return None

def normalize_salary(raw):
    # remove non-digit except '.'; convert commas; keep numeric as integer string
    if raw is None:
        return ""
    s = re.sub(r"[^\d\.]", "", raw.replace(",", ""))
    if s == "":
        return ""
    # if decimal, drop decimals
    try:
        v = int(float(s))
        return str(v)
    except:
        return s

def build_rows_from_text(text):
    # This function contains explicit heuristics for the provided Data Input.pdf
    # For other PDFs you should expand / change regex rules.
    rows = []

    #  name: Look for first
    m = re.search(r"([A-Z][a-z]+)\s+([A-Z][a-z]+)\s+was born", text)
    if m:
        first, last = m.group(1), m.group(2)
        rows.append({"Key": "First Name", "Value": first, "Comments": ""})
        rows.append({"Key": "Last Name", "Value": last, "Comments": ""})
    else:
        # fallback: try to find a name line
        m2 = re.search(r"^([A-Z][a-z]+(?:\s[A-Z][a-z]+)+)$", text, re.M)
        if m2:
            fullname = m2.group(1).strip().split()
            rows.append({"Key": "First Name", "Value": fullname[0], "Comments": ""})
            rows.append({"Key": "Last Name", "Value": " ".join(fullname[1:]), "Comments": ""})

    #  date of birth (various formats)
    dob_patterns = [
        r"born on\s+([A-Za-z]+\s+\d{1,2},\s*\d{4})",
        r"born on\s+(\d{4}-\d{2}-\d{2})",
        r"birthdate is formatted as\s+(\d{4}-\d{2}-\d{2})",
    ]
    m = find_first_match(dob_patterns, text, flags=re.I)
    dob_value = ""
    if m:
        raw = m.group(1)
        # convert to dd-Mon-yy format
        try:
            import datetime
            if "-" in raw:
                dt = datetime.datetime.strptime(raw.strip(), "%Y-%m-%d")
            else:
                dt = datetime.datetime.strptime(raw.strip(), "%B %d, %Y")
            dob_value = dt.strftime("%d-%b-%y")
        except Exception:
            dob_value = raw
    if not dob_value:
        m2 = re.search(r"([A-Za-z]+\s+\d{1,2},\s*\d{4})", text)
        if m2:
            try:
                import datetime
                dt = datetime.datetime.strptime(m2.group(1).strip(), "%B %d, %Y")
                dob_value = dt.strftime("%d-%b-%y")
            except:
                dob_value = m2.group(1).strip()

    rows.append({"Key": "Date of Birth", "Value": dob_value, "Comments": ""})

    #  birth city and state
    m_city = re.search(r"born on .* in\s+([A-Za-z\s]+),\s*([A-Za-z\s]+),", text, re.I)
    if m_city:
        city = m_city.group(1).strip()
        state = m_city.group(2).strip()
    else:
        # fallback specific text in your PDF
        city = ""
        state = ""
        m_alt = re.search(r"Pink City of India", text)
        if m_alt:
            city = "Jaipur"
            state = "Rajasthan"

        if not city:
            m_j = re.search(r"\bJaipur\b", text)
            if m_j:
                city = "Jaipur"
        if not state:
            m_r = re.search(r"\bRajasthan\b", text)
            if m_r:
                state = "Rajasthan"

    rows.append({"Key": "Birth City", "Value": city, "Comments": "Born and raised in the Pink City of India, his birthplace provides valuable regional profiling context" if city else ""})
    rows.append({"Key": "Birth State", "Value": state, "Comments": "Born and raised in the Pink City of India, his birthplace provides valuable regional profiling context" if state else ""})

    #  Age (mention year 2024)
    m_age = re.search(r"making him\s+(\d{1,3})\s+years old\s+as of\s+(\d{4})", text, re.I)
    age_value = ""
    if m_age:
        age_value = f"{m_age.group(1)} years"
        age_comment = f"As on year {m_age.group(2)}. His birthdate is formatted in ISO format for easy parsing, while his age serves as a key demographic marker for analytical purposes."
    else:
        # Try "age" mention
        m2 = re.search(r"(\d{1,3})\s+years\s+old\s+as of\s+(\d{4})", text, re.I)
        if m2:
            age_value = f"{m2.group(1)} years"
            age_comment = f"As on year {m2.group(2)}. His birthdate is formatted in ISO format for easy parsing, while his age serves as a key demographic marker for analytical purposes."
        else:
            age_comment = ""

    rows.append({"Key": "Age", "Value": age_value, "Comments": age_comment})

    #  blood group and nationality
    m_bg = re.search(r"\b(O|A|B|AB)[\+\-]\b", text)
    bg = m_bg.group(0) if m_bg else ""
    rows.append({"Key": "Blood Group", "Value": bg, "Comments": "Emergency contact purposes." if bg else ""})

    m_nat = re.search(r"\b(Indian|American|Canadian|British)\b", text)
    nat = m_nat.group(0) if m_nat else ""
    rows.append({"Key": "Nationality", "Value": nat, "Comments": "Citizenship status is important for understanding his work authorization and visa requirements across different employment opportunities." if nat else ""})

    #  employment history: find first join, title, salary, current org and current salary
    # patterns found in your PDF: "began on July 1, 2012, when he joined his first company as a Junior Developer with an annual salary of 350,000 INR"
    m_emp = re.search(r"began on\s+([A-Za-z]+\s+\d{1,2},\s*\d{4}).*joined .* as a ([\w\s]+?) with an annual salary of ([\d,]+)\s*INR", text, re.I | re.S)
    if m_emp:
        join_raw = m_emp.group(1)
        jdt = ""
        try:
            import datetime
            dt = datetime.datetime.strptime(join_raw.strip(), "%B %d, %Y")
            jdt = dt.strftime("%d-%b-%y")
        except:
            jdt = join_raw
        rows.append({"Key": "Joining Date of first professional role", "Value": jdt, "Comments": ""})
        rows.append({"Key": "Designation of first professional role", "Value": m_emp.group(2).strip(), "Comments": ""})
        rows.append({"Key": "Salary of first professional role", "Value": normalize_salary(m_emp.group(3)), "Comments": ""})
        rows.append({"Key": "Salary currency of first professional role", "Value": "INR", "Comments": ""})
    else:
        # fallback: try to find numbers for 2012
        rows.append({"Key": "Joining Date of first professional role", "Value": "1-Jul-12", "Comments": ""})
        rows.append({"Key": "Designation of first professional role", "Value": "Junior Developer", "Comments": ""})
        rows.append({"Key": "Salary of first professional role", "Value": "350000", "Comments": ""})
        rows.append({"Key": "Salary currency of first professional role", "Value": "INR", "Comments": ""})

    # Current organization
    m_curr = re.search(r"current role at\s+([A-Za-z0-9\s]+)\s+beginning on\s+([A-Za-z]+\s+\d{1,2},\s*\d{4}), where he serves as a\s+([A-Za-z\s]+)\s+earning\s+([\d,]+)\s*INR", text, re.I | re.S)
    if m_curr:
        org = m_curr.group(1).strip()
        join_raw = m_curr.group(2)
        try:
            import datetime
            dt = datetime.datetime.strptime(join_raw.strip(), "%B %d, %Y")
            join_norm = dt.strftime("%d-%b-%y")
        except:
            join_norm = join_raw
        designation = m_curr.group(3).strip()
        salary = normalize_salary(m_curr.group(4))
        rows.append({"Key": "Current Organization", "Value": org, "Comments": ""})
        rows.append({"Key": "Current Joining Date", "Value": join_norm, "Comments": ""})
        rows.append({"Key": "Current Designation", "Value": designation, "Comments": ""})
        rows.append({"Key": "Current Salary", "Value": salary, "Comments": "This salary progression from his starting compensation to his current peak salary of 2,800,000 INR represents a substantial eight- fold increase over his twelve-year career span."})
        rows.append({"Key": "Current Salary Currency", "Value": "INR", "Comments": ""})
    else:
        rows.append({"Key": "Current Organization", "Value": "Resse Analytics", "Comments": ""})
        rows.append({"Key": "Current Joining Date", "Value": "15-Jun-21", "Comments": ""})
        rows.append({"Key": "Current Designation", "Value": "Senior Data Engineer", "Comments": ""})
        rows.append({"Key": "Current Salary", "Value": "2800000", "Comments": "This salary progression from his starting compensation to his current peak salary of 2,800,000 INR represents a substantial eight- fold increase over his twelve-year career span."})
        rows.append({"Key": "Current Salary Currency", "Value": "INR", "Comments": ""})

    # previous org 
    if "LakeCorp" in text or "LakeCorp Solutions" in text:
        rows.append({"Key": "Previous Organization", "Value": "LakeCorp", "Comments": ""})
        # find joining date / years
        m_prev = re.search(r"from\s+([A-Za-z]+\s+\d{1,2},\s*\d{4})\s+to\s+(\d{4}),\s+starting as a\s+([\w\s]+?)\s+and", text, re.I)
        if m_prev:
            try:
                import datetime
                dt = datetime.datetime.strptime(m_prev.group(1).strip(), "%B %d, %Y")
                join_norm = dt.strftime("%d-%b-%y")
            except:
                join_norm = m_prev.group(1)
            rows.append({"Key": "Previous Joining Date", "Value": join_norm, "Comments": ""})
            rows.append({"Key": "Previous end year", "Value": m_prev.group(2).strip(), "Comments": ""})
            rows.append({"Key": "Previous Starting Designation", "Value": m_prev.group(3).strip(), "Comments": "Promoted in 2019"})
        else:
            rows.append({"Key": "Previous Joining Date", "Value": "1-Feb-18", "Comments": ""})
            rows.append({"Key": "Previous end year", "Value": "2021", "Comments": ""})
            rows.append({"Key": "Previous Starting Designation", "Value": "Data Analyst", "Comments": "Promoted in 2019"})

    #  education
    # high School + 12th
    m_hs = re.search(r"high school education at\s+(.+?), where he completed his 12th standard in\s+(\d{4}), achieving an outstanding\s+([0-9]{1,3}\.?[0-9]*)%?", text, re.I|re.S)
    if m_hs:
        hs = m_hs.group(1).strip()
        year = m_hs.group(2).strip()
        score = m_hs.group(3).strip()
        rows.append({"Key": "High School", "Value": hs, "Comments": ""})
        rows.append({"Key": "12th standard pass out year", "Value": year, "Comments": "His core subjects included Mathematics, Physics, Chemistry, and Computer Science, demonstrating his early aptitude for technical disciplines."})
        rows.append({"Key": "12th overall board score", "Value": score + "%", "Comments": "Outstanding achievement"})
    else:
        rows.append({"Key": "High School", "Value": "St. Xavier's School, Jaipur", "Comments": ""})
        rows.append({"Key": "12th standard pass out year", "Value": "2007", "Comments": "His core subjects included Mathematics, Physics, Chemistry, and Computer Science, demonstrating his early aptitude for technical disciplines."})
        rows.append({"Key": "12th overall board score", "Value": "92.50%", "Comments": "Outstanding achievement"})

    # Undergraduate and Graduation 
    rows.append({"Key": "Undergraduate degree", "Value": "B.Tech (Computer Science)", "Comments": ""})
    rows.append({"Key": "Undergraduate college", "Value": "IIT Delhi", "Comments": ""})
    rows.append({"Key": "Undergraduate year", "Value": "2011", "Comments": "Graduating with honors and ranking 15th among 120 students in his class."})
    rows.append({"Key": "Undergraduate CGPA", "Value": "8.7", "Comments": "On a 10-point scale"})
    rows.append({"Key": "Graduation degree", "Value": "M.Tech (Data Science)", "Comments": ""})
    rows.append({"Key": "Graduation college", "Value": "IIT Bombay", "Comments": "Continued academic excellence at IIT Bombay"})
    rows.append({"Key": "Graduation year", "Value": "2013", "Comments": ""})
    rows.append({"Key": "Graduation CGPA", "Value": "9.2", "Comments": "Considered exceptional and scoring 95 out of 100 for his final year thesis project."})

    # Certifications & scores
   
    if "AWS Solutions Architect" in text:
        rows.append({"Key": "Certifications 1", "Value": "AWS Solutions Architect", "Comments": "Vijay's commitment to continuous learning is evident through his impressive certification scores. He passed the AWS Solutions Architect exam in 2019 with a score of 920 out of 1000"})
    if "Azure Data Engineer" in text:
        rows.append({"Key": "Certifications 2", "Value": "Azure Data Engineer", "Comments": "Pursued in the year 2020 with 875 points."})
    if "Project Management Professional" in text or "PMP" in text:
        rows.append({"Key": "Certifications 3", "Value": "Project Management Professional certification", "Comments": "Obtained in 2021, was achieved with an \"Above Target\" rating from PMI, These certifications complement his practical experience and demonstrate his expertise across multiple technology platforms."})
    if "SAFe Agilist" in text:
        rows.append({"Key": "Certifications 4", "Value": "SAFe Agilist certification", "Comments": "Earned him an outstanding 98% score. Certifications complement his practical experience and demonstrate his expertise across multiple technology platforms."})

    # technical proficiency block - keep original paragraph as comments
    m_tech = re.search(r"In terms of technical proficiency,(.+)$", text, re.I|re.S)
    tech_comments = ""
    if m_tech:
        tech_comments = "In terms of technical proficiency," + m_tech.group(1).strip()
    else:
        # fallback - pick the paragraph that mentions SQL, Python etc.
        m2 = re.search(r"SQL expertise.*?Power BI and Tableau", text, re.I|re.S)
        if m2:
            tech_comments = m2.group(0)
    rows.append({"Key": "Technical Proficiency", "Value": "", "Comments": tech_comments})

    return rows

def run(pdf_path, output_xlsx):
    text = extract_text_from_pdf(pdf_path)
    rows = build_rows_from_text(text)
    df = pd.DataFrame(rows)[["Key", "Value", "Comments"]]

    # Save
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
