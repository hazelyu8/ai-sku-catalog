# pipeline.py

import sqlite3
import pytesseract
from PIL import Image
from docx import Document
import openpyxl

# ------------------------
# 1. Extract text
# ------------------------
def extract_text(file_path):
    # (your extract_text function here)
    ...

# ------------------------
# 2. Parse vendor document
# ------------------------
def parse_vendor_doc(text):
    # (your parse_vendor_doc function here)
    ...

# ------------------------
# 3. Save to SQLite
# ------------------------
def save_to_db(fields):
    # (your save_to_db function here)
    ...

# ------------------------
# 4. Export to Excel
# ------------------------
def export_to_excel_template(fields, template_path="data/specs/sample_specs.xlsx"):
    # (your export_to_excel_template function here)
    ...
