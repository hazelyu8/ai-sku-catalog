import os
import glob
import sqlite3
from langchain.agents import initialize_agent, Tool
from langchain_openai import ChatOpenAI

# --- Import your pipeline functions ---
from pipeline import extract_text, parse_vendor_doc, save_to_db, export_to_excel_template

# ------------------------
# 1. Tool wrapper for your pipeline
# ------------------------
def process_vendor_file(file_path: str):
    """Process a vendor doc/image and update DB + Excel."""
    raw_text = extract_text(file_path)
    fields = parse_vendor_doc(raw_text)
    save_to_db(fields)
    export_to_excel_template(fields, "data/specs/sample_specs.xlsx")
    return f"✅ Processed {file_path} and added {fields.get('SKU')} to Excel."

def process_latest_file(folder="data/specs/"):
    """Find latest file in data/specs and process it."""
    files = glob.glob(os.path.join(folder, "*.docx")) + \
            glob.glob(os.path.join(folder, "*.jpg")) + \
            glob.glob(os.path.join(folder, "*.jpeg")) + \
            glob.glob(os.path.join(folder, "*.png"))
    if not files:
        return "❌ No files found in data/specs/"
    latest = max(files, key=os.path.getmtime)
    return process_vendor_file(latest)

def get_last_sku():
    """Fetch the last SKU stored in the database."""
    conn = sqlite3.connect("sku_catalog.db")
    cur = conn.cursor()
    cur.execute("SELECT sku, batch_lot, qty, date FROM sku_catalog ORDER BY id DESC LIMIT 1")
    row = cur.fetchone()
    conn.close()
    if row:
        return f"📦 Last entry → SKU: {row[0]}, Lot: {row[1]}, Qty: {row[2]}, Date: {row[3]}"
    else:
        return "❌ No SKUs found in database."

# ------------------------
# 2. Define tools for the agent
# ------------------------
tools = [
    Tool(
        name="Process Vendor File",
        func=process_vendor_file,
        description="Process a specific vendor doc/image by giving its file path."
    ),
    Tool(
        name="Process Latest File",
        func=process_latest_file,
        description="Process the latest vendor doc/image from data/specs/."
    ),
    Tool(
        name="Get Last SKU",
        func=get_last_sku,
        description="Fetch the last SKU that was added to the database."
    )
]

# ------------------------
# 3. Create the ChatGPT agent
# ------------------------
llm = ChatOpenAI(temperature=0, model="gpt-4o-mini")  # requires OPENAI_API_KEY
agent = initialize_agent(tools, llm, agent="zero-shot-react-description", verbose=True)

# ------------------------
# 4. Chat loop
# ------------------------
if __name__ == "__main__":
    print("🤖 AI SKU Agent is ready! Type 'exit' to quit.\n")
    while True:
        user_input = input("You: ")
        if user_input.lower() in ["exit", "quit"]:
            print("👋 Goodbye!")
            break
        response = agent.run(user_input)
        print("Agent:", response)

