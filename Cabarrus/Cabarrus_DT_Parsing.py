import os
import re
from datetime import datetime
import pytesseract
from PIL import Image
import pdf2image
from google.oauth2.service_account import Credentials
import gspread

# --- Tesseract path ---
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Tesseract\tesseract.exe"

# --- Base & Google Sheet config ---
base_dir = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cabarrus"
SHEET_ID = "1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ"
SHEET_NAME = "Cabarrus County"
CREDENTIALS_FILE = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cabarrus\credentials.json"

# ---- Step 1: Get latest folder ----
def get_latest_scraped_folder(base_path):
    pattern = re.compile(r"Cabarrus Scraped File (\d{2}-\d{2}-\d{4})")
    dated_folders = []
    for folder in os.listdir(base_path):
        match = pattern.match(folder)
        if match:
            try:
                folder_date = datetime.strptime(match.group(1), "%m-%d-%Y")
                dated_folders.append((folder_date, os.path.join(base_path, folder)))
            except:
                continue
    if not dated_folders:
        raise Exception("❌ No valid 'Cabarrus Scraped File' folders found.")
    latest_folder = max(dated_folders, key=lambda x: x[0])[1]
    print(f"📁 Latest folder found: {os.path.basename(latest_folder)}")
    return latest_folder

# ---- Step 2: Generate OCR TXT files ----
def convert_pdfs_to_txt(pdf_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]
    if not pdf_files:
        print("⚠️ No PDF files found.")
        return
    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_folder, pdf_file)
        txt_filename = os.path.splitext(pdf_file)[0] + ".txt"
        txt_path = os.path.join(output_folder, txt_filename)
        try:
            print(f"🔍 Converting: {pdf_file} → {txt_filename}")
            images = pdf2image.convert_from_path(pdf_path, dpi=300)
            full_text = ""
            for img in images:
                full_text += pytesseract.image_to_string(img) + "\n\n"
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(full_text.strip())
            print(f"✅ Saved: {txt_filename}")
        except Exception as e:
            print(f"❌ Failed to convert {pdf_file}: {e}")

# ---- Step 3a: Locate snippet around 'Property Address' ----
def find_property_address_snippet(text, context_chars=300):
    pattern = re.compile(r'property\s+address["\'):]?', re.IGNORECASE)
    match = pattern.search(text)
    if match:
        start = max(0, match.start() - context_chars)
        end = min(len(text), match.end() + context_chars)
        snippet = text[start:end].strip()
        print(f"\n🔎 Found 'Property Address' reference:\n{'-'*60}")
        print(snippet)
        print(f"{'-'*60}\n")
        return snippet
    else:
        print("⚠️ 'Property Address' not found in text.")
        return None

# ---- Step 3b: Address Extraction Patterns ----
def extract_property_address_from_text(text):
    text = text.replace('\xa0', ' ')
    text = text.replace('”', '"').replace('“', '"')
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'\n+', '\n', text)
    text = text.strip()

    cities = r"CONCORD|KANNAPOLIS|MIDLAND|MT\.?\s*PLEASANT|MT PLEASANT|HARRISBURG|CHARLOTTE|DAVIDSON"
    state = r"(?:North\s+Carolina|NC)"
    zip_code = r"\d{5}(?:-\d{4})?"

    patterns = [
        rf"\b\d{{3,6}}\s+[A-Z0-9\s#.'\-]+,\s*(?:{cities})\s*,?\s*{state}\s*{zip_code}\b",
        rf"has the address of\s+([^\n:]+?\b(?:{cities})\b.*?\b{state}\b\s*{zip_code})(?=.*Property\s+Address)",
        rf"which currently has the address of\s+([^\n]+?\b{cities}\b[^\n]*\b{state}\b\s*{zip_code})",
        rf"whose address is\s+([A-Z0-9\s#.'\-]+,\s*[A-Z\s]+,\s*[A-Z]+\s+\d{{5}}(?:-\d{{4}})?)",
        rf"\b\d{{3,6}}\s+[A-Za-z0-9\s#.'\-]+,\s*(?:{cities})\s*,\s*NC\s*{zip_code}\b",
        rf"\b\d{{1,6}}\s+[A-Z0-9\s#.'\-]+,\s*(?:{cities})\s*,\s*NC\s*{zip_code}?\b",
        rf"which currently has the address of\s+([A-Z0-9\s#.'\-]+,\s*(?:{cities})\s*,\s*NC\s*{zip_code})",
    ]

    found = set()
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for match in matches:
            if isinstance(match, tuple):
                match = " ".join(match)
            address = re.sub(r"[\[\]\{\}\(\)<>\"']", "", match).strip()
            address = re.sub(r"\s+", " ", address)
            zip_match = re.search(zip_code, address)
            if zip_match:
                address = address[:zip_match.end()]
            if len(address) < 150:
                found.add(address)

    return list(found)

# ---- Step 3c: Fallback Address Heuristic ----
def fallback_extract_address_from_snippet(snippet):
    lines = snippet.split("\n")
    candidate = ""
    for i, line in enumerate(lines):
        if re.search(r'\b(?:NC|North Carolina)\b', line, re.IGNORECASE) and re.search(r'\d{5}', line):
            before = lines[i - 1].strip() if i - 1 >= 0 else ""
            current = line.strip()
            candidate = f"{before} {current}".strip()
            break

    candidate = re.sub(r'[">«»]', '', candidate)
    candidate = re.sub(r'\s+', ' ', candidate)

    # Remove trailing "(Property Address...)" or variants
    candidate = re.sub(r'\(?["“]?\s*Property Address[^\w]*$', '', candidate, flags=re.IGNORECASE).strip()

    return candidate if len(candidate) > 10 else None


# ---- Step 4: Update Google Sheet ----
def update_sheet_with_addresses(sheet_id, sheet_name, credentials_file, ocr_folder):
    creds = Credentials.from_service_account_file(credentials_file, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id).worksheet(sheet_name)
    data = sheet.get_all_values()
    header = data[0]
    rows = data[1:]

    for i in range(1, 6):
        col = f"Property Address {i}"
        if col not in header:
            header.append(col)
    sheet.update("A1", [header])

    addr_start_col = header.index("Property Address 1") + 1
    pdf_col_index = header.index("Deed of Trust PDF")

    for i, row in enumerate(rows, start=2):
        pdf_filename = row[pdf_col_index] if len(row) > pdf_col_index else ""
        if not pdf_filename or "not found" in pdf_filename.lower():
            continue

        first_addr_index = header.index("Property Address 1")
        if len(row) > first_addr_index and row[first_addr_index].strip():
            continue

        txt_filename = os.path.splitext(pdf_filename)[0] + ".txt"
        txt_path = os.path.join(ocr_folder, txt_filename)

        if not os.path.exists(txt_path):
            print(f"⚠️ OCR file missing: {txt_filename}")
            continue

        with open(txt_path, "r", encoding="utf-8") as f:
            text = f.read()

        print(f"\n📄 Processing: {txt_filename}")
        snippet = find_property_address_snippet(text)
        addresses = extract_property_address_from_text(text)

        if not addresses and snippet:
            fallback = fallback_extract_address_from_snippet(snippet)
            if fallback:
                print(f"🧩 Fallback extracted: {fallback}")
                addresses = [fallback]

        if addresses:
            print(f"✅ Found {len(addresses)} address(es):")
            for j, addr in enumerate(addresses[:5]):
                print(f"   • {addr}")
            for j in range(min(5, len(addresses))):
                sheet.update_cell(i, addr_start_col + j, addresses[j])
        else:
            print("⚠️ No address found.")
            sheet.update_cell(i, addr_start_col, "No Address Found")

    print("\n✅ Done updating sheet with property addresses.")

# ---- Run Script ----
if __name__ == "__main__":
    try:
        latest_folder = get_latest_scraped_folder(base_dir)
        ocr_folder = os.path.join(latest_folder, "OCR Logs")

        if not os.path.exists(ocr_folder) or not os.listdir(ocr_folder):
            print("🛠 OCR Logs folder missing or empty. Generating OCR...")
            convert_pdfs_to_txt(latest_folder, ocr_folder)

        update_sheet_with_addresses(SHEET_ID, SHEET_NAME, CREDENTIALS_FILE, ocr_folder)

    except Exception as e:
        print(f"❌ Script error: {e}")
