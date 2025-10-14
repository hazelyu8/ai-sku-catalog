import streamlit as st
import sqlite3
import pytesseract
from PIL import Image
from docx import Document
import pandas as pd
import openpyxl
import os


# ------------------------
# 1. OCR & DOCX Extraction
# ------------------------
def extract_text_from_image(image_file):
    """Extract text from an uploaded image using OCR."""
    img = Image.open(image_file)
    text = pytesseract.image_to_string(img)
    return text


def extract_text_from_docx(docx_file):
    """Extract both paragraph and table text from DOCX."""
    doc = Document(docx_file)
    text_parts = []

    # Paragraphs
    for para in doc.paragraphs:
        if para.text.strip():
            text_parts.append(para.text.strip())

    # Tables
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if len(cells) == 2:
                text_parts.append(f"{cells[0]} {cells[1]}")
            elif cells:
                text_parts.append(" | ".join(cells))

    return "\n".join(text_parts)


# ------------------------
# 2. Parse Fields
# ------------------------
def parse_vendor_doc(text):
    """Parse structured fields from vendor document text."""
    fields = {
        "Customer": None,
        "Product Description": None,
        "Batch/Lot No.": None,
        "Date": None,
        "SKU": None,
        "Qty": None,
    }

    for line in text.splitlines():
        line = line.strip()
        for key in fields.keys():
            if line.lower().startswith(key.lower()):
                parts = line.split(":", 1)
                if len(parts) == 2:
                    fields[key] = parts[1].strip()
    return fields


# ------------------------
# 3. Save to SQLite (skip duplicates)
# ------------------------
def save_to_db(fields):
    """Save parsed data to local SQLite database — skip duplicates."""
    conn = sqlite3.connect("sku_catalog.db")
    cur = conn.cursor()

    # Create table if it doesn't exist
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

    sku = fields["SKU"]
    batch = fields["Batch/Lot No."]

    # ✅ Check for duplicates (same SKU + Batch/Lot No.)
    cur.execute("""
    SELECT COUNT(*) FROM sku_catalog WHERE sku = ? AND batch_lot = ?
    """, (sku, batch))
    exists = cur.fetchone()[0]

    if exists > 0:
        st.warning(f"⚠️ Entry with SKU '{sku}' and Batch '{batch}' already exists. Skipped saving.")
    else:
        cur.execute("""
        INSERT INTO sku_catalog (customer, product_desc, batch_lot, date, sku, qty)
        VALUES (?, ?, ?, ?, ?, ?)
        """, (
            fields["Customer"],
            fields["Product Description"],
            fields["Batch/Lot No."],
            fields["Date"],
            fields["SKU"],
            fields["Qty"]
        ))
        conn.commit()
        st.success(f"✅ Saved new entry: SKU {sku}, Batch {batch}")

    conn.close()


# ------------------------
# 4. Export full DB → Excel
# ------------------------
def export_db_to_excel(
    db_path="sku_catalog.db",
    excel_path="/Users/hazba/Documents/GitHub/rcos-f25/ai-sku-catalog/data/specs/sample_specs.xlsx",
    sheet_name="Master Sheet - 12th Floor"
):
    """Read all database entries and export them into Excel cleanly."""
    if not os.path.exists(db_path):
        st.error("❌ Database not found. Upload and save a document first.")
        return

    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT * FROM sku_catalog", conn)
    conn.close()

    if df.empty:
        st.warning("⚠️ Database is empty — nothing to export yet.")
        return

    # Remove ID for cleaner Excel layout
    df = df.drop(columns=["id"])

    try:
        # Open the workbook if it exists
        wb = openpyxl.load_workbook(excel_path)
        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(sheet_name)
        else:
            ws = wb[sheet_name]

        # Clear any old data (optional)
        for row in ws["A2:F1000"]:
            for cell in row:
                cell.value = None

        # Write headers
        headers = ["Customer", "Product Description", "Batch/Lot No.", "Date", "SKU", "Qty"]
        for col, header in enumerate(headers, start=1):
            ws.cell(row=1, column=col, value=header)

        # Write database rows starting at row 2
        for row_idx, row_data in enumerate(df.values, start=2):
            for col_idx, value in enumerate(row_data, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)

        wb.save(excel_path)
        st.success(f"✅ Exported {len(df)} rows from database to '{sheet_name}' in Excel!")

    except Exception as e:
        st.error(f"❌ Could not export to Excel. Make sure the file is closed.\n\nError: {e}")


# ------------------------
# 5. Streamlit UI
# ------------------------
st.title("📦 AI-Powered SKU Catalog Agent")
st.markdown("""
Upload a vendor document (.docx or image).  
This app will:
1. Extract key SKU details  
2. Save them in a database (no duplicates!)  
3. Let you export all entries to Excel
""")

uploaded_file = st.file_uploader("Upload a vendor document (JPG, PNG, DOCX)",
                                 type=["jpg", "png", "jpeg", "docx"])

if uploaded_file:
    st.success("✅ File uploaded successfully!")

    if uploaded_file.type in ["image/jpeg", "image/png"]:
        text = extract_text_from_image(uploaded_file)
    else:
        text = extract_text_from_docx(uploaded_file)

    st.subheader("🔎 Extracted Text")
    st.text(text if text.strip() else "(No text detected)")

    fields = parse_vendor_doc(text)
    st.subheader("📦 Parsed Fields")
    st.json(fields)

    if st.button("💾 Save to Database"):
        if any(fields.values()):
            save_to_db(fields)
        else:
            st.warning("⚠️ No valid fields found — check your document.")


# Export button
if st.button("📤 Export Database to Excel"):
    export_db_to_excel()

# Preview DB
if os.path.exists("sku_catalog.db"):
    conn = sqlite3.connect("sku_catalog.db")
    df = pd.read_sql_query("SELECT * FROM sku_catalog", conn)
    conn.close()
    if not df.empty:
        st.subheader("📊 Current Database Entries")
        st.dataframe(df.tail(10))
