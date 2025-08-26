import os
import re
from datetime import datetime
import pytesseract
from PIL import Image
import pdf2image
import gspread
from google.oauth2.service_account import Credentials
from difflib import get_close_matches

# --- Tesseract path ---
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Tesseract\tesseract.exe"
Image.MAX_IMAGE_PIXELS = None  # Disable decompression bomb check

# --- Base config ---
base_dir = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cleveland\Scraped and Downloads"
SHEET_ID = "1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ"
SHEET_NAME = "Cleveland County"
CREDENTIALS_FILE = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cleveland\credentials.json"

# --- Get latest folder ---
def get_latest_scraped_folder(base_path):
    pattern = re.compile(r"Cleveland Scraped File (\d{2}-\d{2}-\d{4})")
    dated = []
    for folder in os.listdir(base_path):
        match = pattern.match(folder)
        if match:
            try:
                date = datetime.strptime(match.group(1), "%m-%d-%Y")
                dated.append((date, os.path.join(base_path, folder)))
            except:
                continue
    if not dated:
        raise Exception("❌ No valid Cleveland folders found.")
    return os.path.join(max(dated)[1], "Deed of Trust Files")

# --- OCR Conversion ---
def convert_pdfs_to_txt(pdf_folder, txt_folder):
    os.makedirs(txt_folder, exist_ok=True)
    converted = []
    for pdf_file in os.listdir(pdf_folder):
        if not pdf_file.lower().endswith(".pdf"):
            continue
        pdf_path = os.path.join(pdf_folder, pdf_file)
        txt_file = os.path.splitext(pdf_file)[0] + ".txt"
        txt_path = os.path.join(txt_folder, txt_file)
        if os.path.exists(txt_path):
            continue
        try:
            print(f"🔍 Converting: {pdf_file}")
            images = pdf2image.convert_from_path(pdf_path, dpi=400, first_page=1, last_page=3)
            text = "\n\n".join([pytesseract.image_to_string(img) for img in images])
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(text.strip())
            print(f"✅ Saved: {txt_file}")
            converted.append((pdf_file, txt_file))
        except Exception as e:
            print(f"❌ Error converting {pdf_file}: {e}")
    return converted

# --- Anchor Scanner ---
def print_property_address_anchor_context(text, context_chars=300):
    anchor_phrases = ["property address", '“property address”', '("property address")', 'property address":', 'property address’;']
    for phrase in anchor_phrases:
        for match in re.finditer(re.escape(phrase), text, flags=re.IGNORECASE):
            start = max(0, match.start() - context_chars)
            end = min(len(text), match.end() + context_chars)
            snippet = text[start:end].replace('\n', ' ')
            print(f"\n🔍 Anchor Detected [{phrase}] — Context:\n{snippet}\n{'-'*80}")

# --- Fuzzy City Cleaner ---
def clean_city_name_fuzzy(address):
    known_cities = ["Shelby", "Kings Mountain", "Cherryville", "Grover", "Boiling Springs", "Lawndale", "Waco", "Fallston", "Charlotte"]
    for city in known_cities:
        matches = get_close_matches(city.lower(), [address.lower()], n=1, cutoff=0.8)
        if matches and city.lower() not in address.lower():
            fuzzy_match = matches[0]
            address = re.sub(fuzzy_match, city, address, flags=re.IGNORECASE)
    return address

# --- Property Address Extractor ---
def extract_property_address_from_text(text):
    cities = r"SHELBY|KINGS MOUNTAIN|CHERRYVILLE|GROVER|BOILING SPRINGS|LAWNDALE|WACO|FALLSTON|CHARLOTTE"
    state = r"(?:North\\s*Carolina|NC)"
    zip_code = r"\d{5}(?:-\d{4})?"

    patterns = [
        rf"\b\d{{1,6}}\s+[A-Z0-9\s#.'\-]+,\s*(?:{cities})\s*,?\s*{state}\s*{zip_code}\b",
        rf"\b\d{{1,6}}\s+[A-Z0-9\s#.'\-]+?\s+(?:{cities})\s+{state}\s*{zip_code}",
        rf"([0-9]{{1,6}} [A-Z0-9\s\-']+)\s+(?:{cities})\s*[, ]*\s*{state}\s*{zip_code}",
        rf"(\d{{1,6}}[^\n]+?(?:{cities})[^\n]*{state}[^\n]*{zip_code})",
        rf"which\s+(?:has|currently\s+has)(?:\s+the)?\s+address\s+of\s+[oa]*\s*(\d{{1,6}}\s+[A-Z0-9\s#.'\-]+?\s+(?:{cities})\b.*?)\s+{state}\s*{zip_code}",
        rf"has the address of\s+([^\n:]+?\b(?:{cities})\b.*?\b{state}\b\s*{zip_code})",
        rf"which currently has the address of\s+([^\n]+?\b{cities}\b[^\n]*\b{state}\b\s*{zip_code})",
        rf"(?:Property Address[;:]*\s*)?TBD\s+[A-Z0-9\s\-#.'\"]+,\s*(?:{cities})\s*,?\s*{state}\s*{zip_code}"
    ]

    found = []
    for idx, pattern in enumerate(patterns):
        try:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                if isinstance(match, tuple):
                    match = " ".join(match)
                cleaned = re.sub(r"[\[\]<>\"'()]", "", match).strip()
                cleaned = re.sub(r'\s+', ' ', cleaned)
                cleaned = clean_city_name_fuzzy(cleaned)
                if len(cleaned) < 150:
                    print(f"🔍 Pattern {idx+1} matched: {cleaned}")
                    found.append(cleaned)
        except Exception as e:
            print(f"⚠️ Pattern {idx+1} error: {e}")
    return found

# --- Snippet Finder ---
def find_property_address_snippet(text, context_chars=300):
    pattern = re.compile(r'property\s+address["\'):]?', re.IGNORECASE)
    match = pattern.search(text)
    if match:
        start = max(0, match.start() - context_chars)
        end = min(len(text), match.end() + context_chars)
        return text[start:end].strip()
    return None

# --- Fallback from snippet ---
def fallback_extract_address_from_snippet(snippet):
    lines = snippet.split("\n")
    for i, line in enumerate(lines):
        if "NC" in line.upper() and re.search(r'\d{5}', line):
            before = lines[i - 1].strip() if i - 1 >= 0 else ""
            return re.sub(r'\s+', ' ', f"{before} {line.strip()}").strip()
    return None

# --- Update Google Sheet ---
def update_sheet(sheet_id, sheet_name, creds_file, txt_folder):
    creds = Credentials.from_service_account_file(creds_file, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id).worksheet(sheet_name)
    data = sheet.get_all_values()
    header, rows = data[0], data[1:]

    # Ensure required headers exist
    required_headers = ["D/T PDF File", "D/T OCR File", "Property Address 1", "Property Address 2", "Property Address 3"]
    for h in required_headers:
        if h not in header:
            header.append(h)
    sheet.update(values=[header], range_name="A1")

    pdf_col = header.index("D/T PDF File")
    ocr_col = header.index("D/T OCR File")
    addr1_col = header.index("Property Address 1")
    addr2_col = header.index("Property Address 2")
    addr3_col = header.index("Property Address 3")

    existing_files = {row[pdf_col].strip(): idx + 2 for idx, row in enumerate(rows) if len(row) > pdf_col and row[pdf_col].strip()}
    txt_files = {f: os.path.join(txt_folder, f) for f in os.listdir(txt_folder) if f.endswith(".txt")}

    for txt_file, txt_path in txt_files.items():
        pdf_file = txt_file.replace(".txt", ".pdf")
        with open(txt_path, "r", encoding="utf-8") as f:
            text = f.read()

        print_property_address_anchor_context(text)

        snippet = find_property_address_snippet(text)
        addresses = extract_property_address_from_text(text)

        if not addresses and snippet:
            fallback = fallback_extract_address_from_snippet(snippet)
            if fallback:
                print(f"🔁 Fallback 1 from snippet: {fallback}")
                addresses.append(fallback)

        if not addresses:
            rough_match = re.search(r'(\d{1,6} [A-Z0-9\s\-#.\']+)\s+(SHELBY|KINGS MOUNTAIN|CHERRYVILLE|CHARLOTTE)[^\n]{0,40}?(?:North Carolina|NC)\s+\d{5}', text, re.IGNORECASE)
            if rough_match:
                second_fallback = " ".join(rough_match.groups()).strip()
                print(f"🔁 Fallback 2 from raw text: {second_fallback}")
                addresses.append(second_fallback)

        # Clean and deduplicate addresses
        cleaned = []
        seen = set()
        for addr in addresses:
            addr = re.sub(r"[“”’']", '', addr)  # Remove weird quotes
            addr = re.sub(r'\s+', ' ', addr).strip()
            addr = addr.title().replace("Nc", "NC")
            addr = clean_city_name_fuzzy(addr)
            if addr and addr.lower() not in seen:
                seen.add(addr.lower())
                cleaned.append(addr)

        addr1 = cleaned[0] if len(cleaned) > 0 else "No Address Found"
        addr2 = cleaned[1] if len(cleaned) > 1 else ""
        addr3 = cleaned[2] if len(cleaned) > 2 else ""

        if pdf_file in existing_files:
            row_num = existing_files[pdf_file]
            sheet.update_cell(row_num, ocr_col + 1, txt_file)
            sheet.update_cell(row_num, addr1_col + 1, addr1)
            sheet.update_cell(row_num, addr2_col + 1, addr2)
            sheet.update_cell(row_num, addr3_col + 1, addr3)
            print(f"✅ Updated row {row_num}: {pdf_file}")
        else:
            new_row = ["" for _ in header]
            new_row[pdf_col] = pdf_file
            new_row[ocr_col] = txt_file
            new_row[addr1_col] = addr1
            new_row[addr2_col] = addr2
            new_row[addr3_col] = addr3
            sheet.append_row(new_row, value_input_option="USER_ENTERED")
            print(f"➕ Appended new row: {pdf_file}")

    print("\n✅ Sheet update complete.")


# --- Run ---
if __name__ == "__main__":
    try:
        dot_folder = get_latest_scraped_folder(base_dir)
        ocr_folder = os.path.join(dot_folder, "OCR Text")
        converted = convert_pdfs_to_txt(dot_folder, ocr_folder)
        update_sheet(SHEET_ID, SHEET_NAME, CREDENTIALS_FILE, ocr_folder)
    except Exception as e:
        print(f"❌ Script error: {e}")
