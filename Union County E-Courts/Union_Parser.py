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

# --- Paths & Google Sheet Info ---
base_dir = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Union\E-Courts"
SHEET_ID = "1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ"
SHEET_NAME = "Union County"
CREDENTIALS_FILE = os.path.join(base_dir, "credentials.json")

# --- Union County Cities ---
union_cities = r"MONROE|STALLINGS|WAXHAW|WEDDINGTON|INDIAN TRAIL|MARSHVILLE|WINGATE|UNIONVILLE|LAKE PARK|FAIRVIEW|MINERAL SPRINGS|HEMBY BRIDGE|WESLEY CHAPEL|CHARLOTTE"
zip_code = r"\d{5}(?:-\d{4})?"
state = r"(?:North\s+Carolina|NC)"

# --- Step 1: Get latest folder ---
def get_latest_scraped_folder(base_path):
    pattern = re.compile(r"Union E-Courts Scraped File (\d{2}-\d{2}-\d{4})")
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
        raise Exception("❌ No valid 'Union E-Courts Scraped File' folders found.")
    latest_folder = max(dated_folders, key=lambda x: x[0])[1]
    print(f"📁 Latest folder found: {os.path.basename(latest_folder)}")
    return latest_folder

# --- Step 2: Generate OCR Logs ---
def convert_pdfs_to_txt(pdf_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]
    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_folder, pdf_file)
        txt_filename = os.path.splitext(pdf_file)[0] + ".txt"
        txt_path = os.path.join(output_folder, txt_filename)
        try:
            print(f"🔍 OCR: {pdf_file}")
            images = pdf2image.convert_from_path(pdf_path, dpi=300)
            text = "\n\n".join([pytesseract.image_to_string(img) for img in images])
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(text.strip())
            print(f"✅ Saved OCR: {txt_filename}")
        except Exception as e:
            print(f"❌ Failed OCR for {pdf_file}: {e}")

# --- Step 3a: Snippet finder ---
def find_property_address_snippet(text, context_chars=300):
    pattern = re.compile(r'property\s+address["\'):]?', re.IGNORECASE)
    match = pattern.search(text)
    if match:
        start = max(0, match.start() - context_chars)
        end = min(len(text), match.end() + context_chars)
        snippet = text[start:end].strip()
        print(f"\n🔎 Snippet around 'Property Address':\n{'-'*60}\n{snippet}\n{'-'*60}")
        return snippet
    return None

# --- Step 3b: Regex-based address extraction ---
def extract_property_address_from_text(text):
    text = re.sub(r'[\xa0“”]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()

    patterns = [
        rf"\b\d{{3,6}}\s+[A-Z0-9\s#.'\-]+(?:,)?\s+(?:{union_cities})\s*,?\s*{state}\s*{zip_code}\b",
        rf"\b\d{{3,6}}\s+[A-Z0-9\s#.'\-]+(?:,)?\s+(?:{union_cities})\s+NC\s*{zip_code}\b",
        rf"\b\d{{1,6}}\s+[A-Z0-9\s#.'\-]+(?:,)?\s+(?:{union_cities})(?!\s+\w+)\s+NC\s+{zip_code}\b"
    ]

    found = set()
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for m in matches:
            address = re.sub(r"[\"'<>]", "", m).strip()
            address = re.sub(r'\s+', ' ', address)
            if len(address) < 150 and not is_courthouse_address(address):
                found.add(address.title())
            else:
                print(f"🚫 Ignored courthouse address: {address}")

    return list(found)

def is_courthouse_address(address):
    """Returns True if the address matches known courthouse patterns."""
    address = address.lower()

    patterns = [
        r"400\s+(n\.?|north)\s+main\s+st\.?(reet)?",   # 400 N. Main St. or 400 North Main Street
        r"union\s+county\s+courthouse",
        r"clerk\s+of\s+(superior\s+)?court",
        r"room\s+1046",
    ]

    city = r"monroe"
    zipcodes = ["28110", "28112"]

    if any(re.search(p, address) for p in patterns) and city in address and any(z in address for z in zipcodes):
        return True
    return False

# --- Step 3c: Fallback extraction ---
def fallback_extract_address_from_snippet(snippet):
    lines = snippet.split("\n")
    for i, line in enumerate(lines):
        if "NC" in line and re.search(zip_code, line):
            before = lines[i - 1].strip() if i > 0 else ""
            combined = f"{before} {line}".strip()
            cleaned = re.sub(r'\s+', ' ', combined)
            return cleaned.title()
    return None

# --- Step 4: Google Sheet Update ---
def update_sheet_with_addresses():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
    data = sheet.get_all_values()
    header = data[0]
    rows = data[1:]

    for i in range(1, 6):
        col = f"Property Address {i}"
        if col not in header:
            header.append(col)
    sheet.update("A1", [header])

    addr_start_col = header.index("Property Address 1") + 1
    case_col_index = header.index("Case Number")

    for i, row in enumerate(rows, start=2):
        case_number = row[case_col_index].strip()
        if not case_number or any(row[addr_start_col + j].strip() for j in range(5) if addr_start_col + j < len(row)):
            continue

        txt_path = os.path.join(ocr_folder, case_number + ".txt")
        if not os.path.exists(txt_path):
            print(f"⚠️ OCR missing for: {case_number}")
            continue

        with open(txt_path, "r", encoding="utf-8") as f:
            text = f.read()

        print(f"\n📄 Processing OCR: {case_number}")
        snippet = find_property_address_snippet(text)
        addresses = extract_property_address_from_text(text)

        if not addresses and snippet:
            fallback = fallback_extract_address_from_snippet(snippet)
            if fallback:
                print(f"🧩 Fallback extracted: {fallback}")
                addresses = [fallback]

        if addresses:
            for j in range(min(5, len(addresses))):
                sheet.update_cell(i, addr_start_col + j, addresses[j])
        else:
            sheet.update_cell(i, addr_start_col, "No Address Found")
            print("⚠️ No address found.")

    print("\n✅ Finished updating Union County sheet.")

# --- Main runner ---
if __name__ == "__main__":
    try:
        latest_folder = get_latest_scraped_folder(base_dir)
        ocr_folder = os.path.join(latest_folder, "OCR Logs")
        if not os.path.exists(ocr_folder) or not os.listdir(ocr_folder):
            print("🛠 OCR folder missing or empty. Generating...")
            convert_pdfs_to_txt(latest_folder, ocr_folder)
        update_sheet_with_addresses()
    except Exception as e:
        print(f"❌ Script error: {e}")
