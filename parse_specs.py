import sqlite3
import pytesseract
from PIL import Image
import re
import pandas as pd
from docx import Document
import os

# ------------------------
# 1. Extract text (image OR docx)
# ------------------------
def extract_text(file_path):
    ext = os.path.splitext(file_path)[1].lower()

    if ext in [".jpg", ".jpeg", ".png"]:
        img = Image.open(file_path)
        text = pytesseract.image_to_string(img)
        return text

    elif ext == ".docx":
        doc = Document(file_path)
        text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        return text

    else:
        raise ValueError(f"Unsupported file type: {ext}")

# ------------------------
# 2. Parse vendor document
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

    # Regex with re.MULTILINE and case-insensitive
    cust_match = re.search(r"Customer[:\-]?\s*(.+)", text, re.MULTILINE | re.IGNORECASE)
    prod_match = re.search(r"(Product Description|Description)[:\-]?\s*(.+)", text, re.MULTILINE | re.IGNORECASE)
    lot_match = re.search(r"(Batch/Lot No\.?|Lot)[:\-]?\s*(.+)", text, re.MULTILINE | re.IGNORECASE)
    date_match = re.search(r"Date[:\-]?\s*(.+)", text, re.MULTILINE | re.IGNORECASE)
    sku_match = re.search(r"SKU[:\-]?\s*(.+)", text, re.MULTILINE | re.IGNORECASE)
    qty_match = re.search(r"(Qty|Quantity)[:\-]?\s*(\d+)", text, re.MULTILINE | re.IGNORECASE)

    if cust_match: fields["Customer"] = cust_match.group(1).strip()
    if prod_match: fields["Product Description"] = prod_match.group(2).strip()
    if lot_match: fields["Batch/Lot No."] = lot_match.group(2).strip()
    if date_match: fields["Date"] = date_match.group(1).strip()
    if sku_match: fields["SKU"] = sku_match.group(1).strip()
    if qty_match: fields["Qty"] = qty_match.group(2).strip()

    return fields

# ------------------------
# 3. Save to SQLite (with duplicate check)
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
        sku TEXT UNIQUE,
        qty INTEGER
    )
    """)

    try:
        cur.execute("""
        INSERT INTO sku_catalog (customer, product_desc, batch_lot, date, sku, qty)
        VALUES (?, ?, ?, ?, ?, ?)
        """, (
            fields["Customer"], fields["Product Description"], fields["Batch/Lot No."],
            fields["Date"], fields["SKU"], fields["Qty"]
        ))
        conn.commit()
        print("✅ Entry saved to database.")
    except sqlite3.IntegrityError:
        print("⚠️ Duplicate entry skipped (same SKU already exists).")

    conn.close()

# ------------------------
# 4. Export DB → Excel + Preview
# ------------------------
def export_to_excel():
    conn = sqlite3.connect("sku_catalog.db")
    df = pd.read_sql_query("SELECT * FROM sku_catalog", conn)
    df.to_excel("sku_catalog.xlsx", index=False)
    conn.close()
    print("📊 Excel export updated: sku_catalog.xlsx")

    # Show preview in terminal
    print("\n🔎 Current Catalog Preview:")
    print(df.to_string(index=False))

# ------------------------
# 5. Run the agent
# ------------------------
if __name__ == "__main__":
    # Swap this path to test with either .jpg/.png or .docx
    file_path = "data/specs/smolder_eyes_document.docx"

    # Step 1: Extract text
    raw_text = extract_text(file_path)
    print("🔎 Extracted Text:\n", raw_text)

    # Step 2: Parse fields
    fields = parse_vendor_doc(raw_text)
    print("📦 Parsed Fields:", fields)

    # Step 3: Save to DB
    save_to_db(fields)

    # Step 4: Update Excel + preview
    export_to_excel()
