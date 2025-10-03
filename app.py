import streamlit as st
import sqlite3
import pytesseract
from PIL import Image
from docx import Document
import re
import pandas as pd
import openpyxl
import os

# ------------------------
# 1. OCR & Docx extraction
# ------------------------
def extract_text_from_image(image_file):
    img = Image.open(image_file)
    text = pytesseract.image_to_string(img)
    return text

def extract_text_from_docx(docx_file):
    doc = Document(docx_file)
    text = "\n".join([para.text for para in doc.paragraphs])
    return text

# ------------------------
# 2. Parse into fields
# ------------------------
def parse_vendor_doc(text):
    fields = {
        "Customer": None,
        "Product Description": None,
        "Batch/Lot No.": None,
        "Date": None,
        "SKU": None,
        "Qty": None
    }

    patterns = {
        "Customer": r"Customer[:\-]?\s*([A-Za-z0-9 ]+)",
        "Product Description": r"(Product Description|Description)[:\-]?\s*([A-Za-z0-9 \-\_]+)",
        "Batch/Lot No.": r"(Batch|Lot|Batch/Lot No.)[:\-]?\s*([A-Za-z0-9\-\_]+)",
        "Date": r"Date[:\-]?\s*([0-9/]+)",
        "SKU": r"SKU[:\-]?\s*([A-Za-z0-9\-\_]+)",
        "Qty": r"(Qty|Quantity)[:\-]?\s*([0-9]+)"
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            fields[key] = match.group(1) if len(match.groups()) == 1 else match.group(2)

    return fields

# ------------------------
# 3. Save to SQLite
# ------------------------
def save_to_db(fields):
    conn = sqlite3.connect("sku_catalog.db")
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS sku_catalog (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        customer TEXT,
        product_desc TEXT,
        batch_lot TEXT,
        date TEXT,
        sku TEXT,
        qty TEXT
    )
    """)
    cur.execute("""
    INSERT INTO sku_catalog (customer, product_desc, batch_lot, date, sku, qty)
    VALUES (?, ?, ?, ?, ?, ?)
    """, (
        fields["Customer"], fields["Product Description"], fields["Batch/Lot No."],
        fields["Date"], fields["SKU"], fields["Qty"]
    ))
    conn.commit()
    conn.close()

# ------------------------
# 4. Save to Excel (template style)
# ------------------------
def save_to_excel(fields, excel_path="data/specs/sample_specs.xlsx", sheet_name="Master Sheet - 12th Floor"):
    if not os.path.exists(excel_path):
        # Create if not exists
        df = pd.DataFrame(columns=["Customer", "Product Description", "Batch/Lot No.", "Date", "SKU", "Qty"])
        df.to_excel(excel_path, index=False, sheet_name=sheet_name)

    wb = openpyxl.load_workbook(excel_path)
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        ws.append(["Customer", "Product Description", "Batch/Lot No.", "Date", "SKU", "Qty"])
    else:
        ws = wb[sheet_name]

    ws.append([
        fields["Customer"], fields["Product Description"], fields["Batch/Lot No."],
        fields["Date"], fields["SKU"], fields["Qty"]
    ])

    wb.save(excel_path)

# ------------------------
# 5. Streamlit UI
# ------------------------
st.title("📦 AI-Powered SKU Catalog Agent")

uploaded_file = st.file_uploader("Upload a vendor document (JPG, PNG, DOCX)", type=["jpg", "png", "jpeg", "docx"])

if uploaded_file:
    st.write("✅ File uploaded!")
    
    if uploaded_file.type in ["image/jpeg", "image/png"]:
        text = extract_text_from_image(uploaded_file)
    else:
        text = extract_text_from_docx(uploaded_file)

    st.subheader("🔎 Extracted Text")
    st.text(text)

    fields = parse_vendor_doc(text)
    st.subheader("📦 Parsed Fields")
    st.json(fields)

    if st.button("Save to Database & Excel"):
        save_to_db(fields)
        save_to_excel(fields)
        st.success("✅ Entry saved to database and Excel.")
