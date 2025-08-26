import os
import glob
import re
import pytesseract
import pdf2image
import gspread
from PIL import Image
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

# --- Config ---
TESSERACT_PATH = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Tesseract\tesseract.exe"
BASE_DIR = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Gaston\Scraped File"
CREDENTIALS_FILE = "credentials.json"
SHEET_KEY = "1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ"
SHEET_TAB = "Gaston County"

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH


# --- Utilities ---
def get_latest_folder(path):
    folders = [os.path.join(path, f) for f in os.listdir(path) if os.path.isdir(os.path.join(path, f))]
    if not folders:
        raise Exception("No subfolders found.")
    return max(folders, key=os.path.getmtime)

def mark_llc_in_sheet(sheet):
    data = sheet.get_all_values()
    headers = data[0]
    rows = data[1:]

    try:
        grantors_col = headers.index("Grantors")
        llc_col = headers.index("LLC Identifier")
    except ValueError:
        print("⚠️ 'Grantors' or 'LLC Identifier' column not found in sheet headers.")
        return set()

    llc_rows = set()
    updates = []

    for i, row in enumerate(rows):
        grantor = row[grantors_col].strip().upper()
        current_llc_value = row[llc_col].strip() if llc_col < len(row) else ""

        # ✅ Skip if already filled
        if current_llc_value:
            continue

        is_llc = "LLC" in grantor
        llc_rows.add(i) if is_llc else None
        cell_value = "True" if is_llc else "False"
        updates.append(gspread.cell.Cell(i + 2, llc_col + 1, cell_value))

    if updates:
        sheet.update_cells(updates)
        print(f"✅ LLC Identifier column updated for {len(updates)} new rows.")
    else:
        print("⏭️ No new LLC Identifier updates needed.")

    return llc_rows

def generate_ocr_logs(pdf_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')]

    for filename in pdf_files:
        try:
            base_name = os.path.splitext(filename)[0]
            output_path = os.path.join(output_folder, f"{base_name}.txt")
            if os.path.exists(output_path):
                print(f"⏭️ Skipping {filename} - OCR log already exists.")
                continue

            pdf_path = os.path.join(pdf_folder, filename)
            images = pdf2image.convert_from_path(pdf_path, dpi=300)
            full_text = ""
            for img in images:
                text = pytesseract.image_to_string(img)
                full_text += text + "\n\n"

            with open(output_path, "w", encoding="utf-8") as f:
                f.write(full_text.strip())
            print(f"✅ OCR log created: {output_path}")
        except Exception as e:
            print(f"⚠️ Error processing {filename}: {e}")


def normalize_raw_address_text(raw):
    raw = re.sub(r'[\{\}\[\]\<\>].*?[\}\]]', '', raw)
    replacements = {
        'Gity': 'City', '28°33': '28033', 'aro': 'NC', 'aro.': 'NC.', 'Nok Carolina': 'North Carolina',
        'North CNClina': 'North Carolina', 'North Carclina': 'North Carolina', 'North Cazolina': 'North Carolina',
        'WeSTtT': 'West', 'HEAL WAY': 'Healthy Way', 'SIBNNETT WRAIL DR': 'Sibbett Trail Dr',
        'IQQE THIRD AVE': '1000 Third Avenue', ' ,': ',', ' , ': ', ','WeSTtT': 'West',
    }
    for wrong, right in replacements.items():
        raw = raw.replace(wrong, right)
    raw = re.sub(r'(?i)^PAGE\s+\d+\s+IN\s+THE\s+GASTON\s+COUNTY\s+PUBLIC\s+REGISTRY\.\s+', '', raw)
    raw = re.sub(r'(?i)(Parcel ID Number:\s*\d+\s*)?(which\s+currently\s+)?has the address of\s+', '', raw)
    return re.sub(r'\s+', ' ', raw).strip('.,"- ')


def find_property_address_snippet(text, context_chars=300):
    pattern = re.compile(r'property\s+address["\'):]?', re.IGNORECASE)
    match = pattern.search(text)
    if match:
        start = max(0, match.start() - context_chars)
        end = min(len(text), match.end() + context_chars)
        return text[start:end].strip()
    return None


def fallback_extract_address_from_snippet(snippet):
    lines = snippet.split("\n")
    for i, line in enumerate(lines):
        if re.search(r'\b(?:NC|North Carolina)\b', line, re.IGNORECASE) and re.search(r'\d{5}', line):
            before = lines[i - 1].strip() if i - 1 >= 0 else ""
            current = line.strip()
            combined = f"{before} {current}".strip()
            combined = normalize_raw_address_text(combined)
            if len(combined) >= 10 and not combined.isupper():
                return combined
    return None


def extract_property_address_from_text(text, filename=""):
    gaston_cities = r"GASTONIA|BELMONT|BESSEMER CITY|CHERRYVILLE|CROWDERS|DALLAS|DELVIEW|HIGH SHOALS|KING'S MOUNTAIN|LOWELL|MCADENVILLE|MOUNT HOLLY|RANLO|SPENCER MOUNTAIN|STANLEY"
    state = r"(?:North\s+Carolina|NC)"
    zip_code = r"\d{5}(?:-\d{4})?"

    patterns = [
        rf"\b\d{{3,6}}\s+[A-Z0-9\s#.'\-]+,\s*(?:{gaston_cities})\s*,?\s*{state}\s*{zip_code}\b",
        rf"has the address of\s+([^\n:]+?\b(?:{gaston_cities})\b.*?\b{state}\b\s*{zip_code})",
        rf"which currently has the address of\s+([^\n]+?\b{gaston_cities}\b[^\n]*\b{state}\b\s*{zip_code})",
        rf"whose address is\s+([A-Z0-9\s#.'\-]+,\s*[A-Z\s]+,\s*[A-Z]+\s+\d{{5}}(?:-\d{{4}})?)",
        rf"\b\d{{3,6}}\s+[A-Z0-9\s#.'\-]+[\[\(]Street[\]\)]\s*,?\s*[A-Z\s]+[aA]?\s*{state}\s*{zip_code}\b",
        rf"\b\d{{3,6}}\s+[A-Z0-9\s#.'\-]+,\s*(?:{gaston_cities})\s*,?\s*{state}\s*,?\s*{zip_code}\b",
        rf"\[Property Address\][\s‘'\":-]*([\dA-Z\s#.'\-]+,\s*(?:{gaston_cities})\s*,?\s*{state}\s*,?\s*{zip_code})",
        rf"commonly known as\s+([^\n,]+,\s*(?:{gaston_cities})\s*,?\s*{state}\s*{zip_code})",
        rf"has the address of\s*([\dA-Z\s#.'\-]+)\s*\[Street\][\s,]*([A-Z\s]+)\s*\[City\][\s,]*{state}.*?{zip_code}",
        
    ]

    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for match in matches:
            match = " ".join(match) if isinstance(match, tuple) else match
            address = re.sub(r"[\[\]\{\}\(\)<>\"']", "", match).strip()
            address = re.sub(r'\s+', ' ', address)
            if len(address) < 10 or address.isdigit():
                continue
            if not re.search(rf"\b(?:{gaston_cities})\b", address, re.IGNORECASE):
                continue
            if not re.search(rf"\b{state}\b", address, re.IGNORECASE):
                continue
            zip_match = re.search(zip_code, address)
            if zip_match:
                address = address[:zip_match.end()]
            address = normalize_raw_address_text(address)
            cleaned = clean_address(address)
            if cleaned and len(cleaned) < 150:
                print(f"📄 [MAIN MATCH] {filename}: {cleaned}")
                return cleaned

    # 🆕 Multi-line fallback: line breaks + "has the address of"
    multi_line_match = re.search(
        rf"has the address of\s*\n*([\dA-Z\s#.'\-]+)[\s\n,]*\b({gaston_cities})\b[\s,]*{state}[\s,]*{zip_code}",
        text,
        re.IGNORECASE
    )
    if multi_line_match:
        address = " ".join(multi_line_match.groups())
        cleaned = clean_address(normalize_raw_address_text(address))
        if cleaned:
            print(f"📄 [MULTILINE MATCH] {filename}: {cleaned}")
            return cleaned

    # 🆕 New pattern: Property Address: 123 Street, City, NC
    tail_text = text[-1000:]
    match = re.search(
        rf"Property Address[:\s]*([^\n,]+,\s*(?:{gaston_cities})\s*,?\s*{state}.*?)\b",
        tail_text,
        re.IGNORECASE
    )
    if match:
        raw_address = match.group(1)
        cleaned = clean_address(normalize_raw_address_text(raw_address))
        if cleaned:
            print(f"📄 [TAIL MATCH] {filename}: {cleaned}")
            return cleaned

    return None


def extract_parcel_id(text, filename=""):
    patterns = [
        r"ID#\s*(\d{6,})",
        r"Parcel ID Number:\s*(\d+)",
        r"Parcel #:\s*(\d+)",
        r"Tax parcel number:\s*(\d+)",
        r"TAX\s+MAP\s+OR\s+PARCEL\s+ID\s+NO\.?:\s*(\d+)",
        r"TAX\s+MAP\s+OR\s+PARCEL\s+10\s+NO\.?:\s*(\d+)",
    ]

    # 🔍 Try full text first
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            parcel = match.group(1).strip()
            print(f"📄 [PARCEL MATCH - FULL] {filename}: {parcel}")
            return parcel

    # 🔍 Fallback: scan last ~1000 characters of text
    tail = text[-1000:]
    for pattern in patterns:
        match = re.search(pattern, tail, re.IGNORECASE)
        if match:
            parcel = match.group(1).strip()
            print(f"📄 [PARCEL MATCH - TAIL] {filename}: {parcel}")
            return parcel

    return None

def clean_address(address):
    address = re.sub(r'\s+', ' ', address.replace('\n', ' ').replace('\r', ' ')).strip()
    address = re.sub(r'^[,.\s]+|[,.\s]+$', '', address)
    address = re.sub(r'[\[\{(]?(Street|City|Zip\s*Code)[\]\})]?', '', address, flags=re.IGNORECASE)
    address = re.sub(r'\bProperty Address\b[:\-]?', '', address, flags=re.IGNORECASE).strip()
    address = address.replace('("Property Address")', '')
    address = ' '.join(word.capitalize() if not word.isupper() else word for word in address.split())
    if re.fullmatch(r'\d{5}(-\d{4})?', address) or address.lower() in ['gastonia', 'charlotte']:
        return None
    return address


def extract_address_and_parcel_id(ocr_logs_folder):
    results = []
    txt_files = glob.glob(os.path.join(ocr_logs_folder, "*.txt"))

    for txt_file in txt_files:
        filename = os.path.basename(txt_file)
        with open(txt_file, 'r', encoding='utf-8', errors='ignore') as f:
            text = f.read()

        snippet = find_property_address_snippet(text)
        address = extract_property_address_from_text(text, filename)
        if not address and snippet:
            address = fallback_extract_address_from_snippet(snippet)

        parcel_id = extract_parcel_id(text, filename)
        results.append({
            "File": filename,
            "Property Address": address or "Not Found",
            "Parcel ID": parcel_id or "Not Found",
        })

    return results


def update_sheet_with_extracted_data(sheet, extracted_data, latest_folder):
    data = sheet.get_all_values()
    headers = data[0]
    rows = data[1:]

    pdf_col = headers.index("Downloaded PDF")
    parcel_col = headers.index("Parcel ID")
    address_col = headers.index("Property Address")

    latest_pdf_names = {os.path.basename(f) for f in glob.glob(os.path.join(latest_folder, "*.pdf"))}
    updates = []

    for i, row in enumerate(rows):
        pdf_name = row[pdf_col].strip()
        if not pdf_name or pdf_name not in latest_pdf_names:
            continue

        txt_name = pdf_name.replace(".pdf", ".txt")
        match = next((entry for entry in extracted_data if entry['File'] == txt_name), None)

        if match:
            row_index = i + 2
            updates.append(gspread.cell.Cell(row_index, parcel_col + 1, match['Parcel ID']))
            updates.append(gspread.cell.Cell(row_index, address_col + 1, match['Property Address']))

    if updates:
        sheet.update_cells(updates)
        print(f"✅ {len(updates)//2} rows updated.")
    else:
        print("⚠️ No updates to apply.")


# --- Main ---
def main():
    print("🚀 Starting Gaston OCR + Sheet Update")
    latest_folder = get_latest_folder(BASE_DIR)
    ocr_logs_folder = os.path.join(latest_folder, "OCR Logs")

    print(f"📂 Latest folder: {latest_folder}")
    generate_ocr_logs(latest_folder, ocr_logs_folder)
    extracted_data = extract_address_and_parcel_id(ocr_logs_folder)

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        CREDENTIALS_FILE, ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    )
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SHEET_KEY).worksheet(SHEET_TAB)

    # ✅ Ensure this line is added
    mark_llc_in_sheet(sheet)

    update_sheet_with_extracted_data(sheet, extracted_data, latest_folder)
    print("✅ Gaston County sheet update complete.")


if __name__ == "__main__":
    main()