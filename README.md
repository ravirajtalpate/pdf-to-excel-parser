# AI-Powered Document Structuring & Data Extraction

Transform unstructured PDF documents into clean, structured Excel files.

Overview :

This project automatically converts an unstructured PDF document into a structured Excel output by detecting key:value pairs, extracting contextual information, and preserving all content.
It was developed as part of an AI internship assignment to demonstrate automated document parsing and data organization using Python.

Features :

Extracts raw text from any PDF using pdfplumber

Identifies key: value pairs using multiple regex-based heuristics

Preserves and outputs 100% of PDF content

Adds contextual comments derived from identified text segments

Generates a clean, structured Output.xlsx file

Can be extended to support multiple document formats



Includes optional demo interface (Streamlit)

Project Structure
.
├── extractor.py            # Main script that parses PDFs and generates Excel
├── Output.xlsx             # Generated sample output
├── requirements.txt        # Dependencies for the project
└── README.md               # Documentation





How It Works
1. Text Extraction

The script uses pdfplumber to extract text from each page of the PDF.

2. Pattern Recognition (Regex-Based)

Multiple-layered regex rules identify:

Personal details

Birth information

Dates

Education

Work experience

Technical skills

Certifications

Fallback rules handle variations in formatting.




3. Structuring the Data

Each extracted element is stored as:

Key

Value

Comments

Everything is appended to a pandas DataFrame.




4. Excel Generation

The final structured table is exported as Output.xlsx using pandas.

Installation

Clone the project:

git clone https://github.com/your-repo/document-extractor.git
cd document-extractor


Install requirements:

pip install -r requirements.txt

Usage

Run the extraction script:

python extractor.py --pdf DataInput.pdf --out Output.xlsx


Arguments:

Flag	Description
--pdf	Path to the input PDF file
--out	Output Excel file path
Output Format



The generated Excel file contains three columns:

Key	Value	Comments

This format matches the evaluation template.

Live Demo (Optional)

A Streamlit demo can be added for hosting on Streamlit Cloud or Render.



Example usage:

streamlit run app.py


This allows users to upload a PDF and download the generated Excel file instantly

Assumptions

PDF follows a semi-structured narrative style

Regex heuristics will be expanded as needed

No AI/ML model is required unless extended

Future Improvements

Add NLP-based context extraction

Add layout-aware PDF parsing

Integrate a hosted API endpoint

Support Word documents and scanned PDFs

License

This project is for assessment and educational purposes only.


