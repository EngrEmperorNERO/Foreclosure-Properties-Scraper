import os
import re
import shutil
import pytesseract
import pandas as pd
from PIL import Image
import pdf2image
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Set Tesseract path
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Tesseract\tesseract.exe"

# ---- Step 1: Convert PDFs to OCR Logs ----
def generate_ocr_logs_from_pdfs(pdf_folder, ocr_folder):
    if os.path.exists(ocr_folder):
        shutil.rmtree(ocr_folder)
    os.makedirs(ocr_folder)

    pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]
    for filename in pdf_files:
        pdf_path = os.path.join(pdf_folder, filename)
        try:
            images = pdf2image.convert_from_path(pdf_path, dpi=300)
            full_text = ""
            for img in images:
                text = pytesseract.image_to_string(img)
                full_text += " " + re.sub(r'\s+', ' ', text).strip()
            txt_filename = os.path.splitext(filename)[0] + "_OCR.txt"
            txt_path = os.path.join(ocr_folder, txt_filename)
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(full_text)
        except Exception as e:
            print(f"❌ Failed to OCR {filename}: {e}")

# ---- Step 2: Extract Address from Text ----
def extract_property_address_from_text(text):
    cities = r"CHARLOTTE|CORNELIUS|DAVIDSON|HUNTERSVILLE|MATTHEWS|MINT HILL|PINEVILLE"
    state = r"(?:North\s+Carolina|NC)"
    zip_code = r"\d{5}(?:-\d{4})?"

    patterns = [
        rf"has the address of\s+([0-9]{{3,6}}[^\n:,]*?\b(?:{cities})\b[^\n:,]*?\b{state}\b[^\d\n]*?{zip_code})",
        rf"\b\d{{3,6}}\s+[A-Z0-9\s#.'\-]+,\s*(?:{cities})\s*,?\s*{state}\s*{zip_code}\b",
        rf"has the address of\s+([\dA-Z\s#.,\-]+?\b{cities}\b\s*»\s*{state}\s*{zip_code})(?=\s*\(\s*\"Property Address\")",
        rf"which (?:currently\s+)?has the address of\s+([\dA-Z\s#.,\-]+?\b(?:{cities})\b[\s,]+{state}[\s\d\-]*)(?=\s*[\(\[].*Property Address)",
        rf"which (?:currently\s+)?has the address of\s+([\dA-Z\s#.,\-]+?\b(?:{cities})\b[\s,]+{state}[\s\d\-]*)(?=\s+and\s+Parcel\s+No)",
        rf"has the address of\s+([\dA-Z\s#.,\-]+?\b(?:{cities})\b[\s,]+{state}[\s\d\-]*)",
        rf"has the address of\s+([\dA-Z\s#.,\-]+?\b{cities}\b\s*(?:[»>:\-]?\s*)?{state}\s*{zip_code})(?=\s*[\(\[\{{]?\s*\"?Property Address\"?)",
        # Garbage-tolerant
        rf"has the address of\s+([^\n:]+?\b(?:{cities})\b.*?\b{state}\b\s*{zip_code})(?=.*Property\s+Address)",
        rf"\b\d{{3,6}}\s+[A-Za-z0-9\s#.'\-]+,\s*(?:{cities})\s*,\s*NC\s*{zip_code}\b",
        rf"whose address is\s+([A-Z0-9\s#.'\-]+,\s*[A-Z\s]+,\s*[A-Z]+\s+\d{{5}}(?:-\d{{4}})?)",
        rf"which (?:currently\s+)?has the address of\s+([\dA-Z\s#.,\-]+?\b(?:{cities})\b[^\n:,]*?\b{state}\b[^\d\n]*?{zip_code})(?=.*Property\s*Address)",
        rf"\b\d{{3,6}}\s+[A-Za-z0-9\s#.'\-]+\s*,\s*{cities}\s*,\s*{state}\s*{zip_code}\s*-\s*\(\s*\"Property Address\"\s*\)",
        rf"which (?:currently\s+)?has the address of\s+([\dA-Z\s#.,\-]+?\b(?:{cities})\b[^\n:,]*?\b{state}\b[^\d\n]*?{zip_code})\s*[\(\[\{{]?[\"'“”‘’]?\s*Property Address[\"'“”‘’]?\s*[\)\]\}}]?",
        rf"which (?:currently\s+)?has the address of\s+([\dA-Z\s#.,\-]+?\b{cities}\b[^\n:,]*?\b{state}\b[^\d\n]*?{zip_code})\s*[\"'“”‘’]?\s*\(?Property Address\)?[\"'“”‘’]?",
        rf"which (?:currently\s+)?has the address of\s+([\dA-Z\s#.,\-]+?\b{cities}\b[^\n:,]*?\b{state}\b[^\d\n]*?{zip_code})\s*\(\s*\"Property Address\"\s*\)",
        rf"which (?:currently\s+)?has the address of\s+([\dA-Z\s#.,\-]+?\b{cities}\b[^\n:,]*?\[Street\][^\n:,]*?\[City\][^\n:,]*?\b{state}\b[^\d\n]*?{zip_code})(?=.*Property\s*Address)",
    ]
    
    found = set()
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for match in matches:
            # Normalize whitespace and remove junk
            address = re.sub(r"[\[\]\{\}\(\)<>]", "", match).strip()
            address = re.sub(r"\s+", " ", address)
            # Require at least 3 words (e.g., "10107 Daufuskie Dr Charlotte")
            if len(address.split()) < 3:
                continue

            # Truncate after ZIP to avoid capturing paragraphs
            zip_match = re.search(zip_code, address)
            if zip_match:
                end = zip_match.end()
                address = address[:end]

            # Remove overly long fragments
            if len(address) < 150:
                found.add(address)
    return list(found)

# ---- Step 3: Update Google Sheet ----
def update_sheet_with_addresses(sheet_id, sheet_name, credentials_file, ocr_folder):
    creds = Credentials.from_service_account_file(credentials_file, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gspread_client = gspread.authorize(creds)
    sheet_service = build('sheets', 'v4', credentials=creds)
    sheet = sheet_service.spreadsheets()

    worksheet = gspread_client.open_by_key(sheet_id).worksheet(sheet_name)
    data = worksheet.get_all_values()
    header = data[0]
    rows = data[1:]

    # Ensure necessary columns exist
    if "OCR Text Filename" not in header:
        header.append("OCR Text Filename")
    for i in range(1, 6):
        col_name = f"Property Address {i}"
        if col_name not in header:
            header.append(col_name)
    header_map = {col: idx for idx, col in enumerate(header)}
    worksheet.update(values=[header], range_name='A1')

    updates = []

    for i, row in enumerate(rows, start=2):
        pdf_filename = row[header_map["PDF Filename"]] if len(row) > header_map["PDF Filename"] else ""
        if not pdf_filename:
            continue

        ocr_filename = os.path.splitext(pdf_filename)[0] + "_OCR.txt"
        ocr_path = os.path.join(ocr_folder, ocr_filename)

        # Skip if already filled
        ocr_cell = row[header_map["OCR Text Filename"]] if len(row) > header_map["OCR Text Filename"] else ""
        if ocr_cell.strip():
            continue

        if not os.path.exists(ocr_path):
            continue

        with open(ocr_path, "r", encoding="utf-8") as file:
            text = file.read()

        # 🔍 Debug OCR snippet
        if "has the address of" in text.lower():
            snippet = re.search(r"has the address of.{0,300}", text, re.IGNORECASE)
            print(f"🧾 Found snippet in {ocr_filename}:")
            print(snippet.group(0) if snippet else "None")

        # 📌 Extract & Print Addresses
        addresses = extract_property_address_from_text(text)
        print(f"🏠 Addresses found in {ocr_filename}:")
        if addresses:
            for a in addresses:
                print("   •", a)
        else:
            print("   ⚠️ No address found")

        row_updates = [("OCR Text Filename", ocr_filename)]
        if addresses:
            for j in range(min(5, len(addresses))):
                row_updates.append((f"Property Address {j+1}", addresses[j]))
        else:
            row_updates.append(("Property Address 1", "No Address Found"))

        for col_name, value in row_updates:
            col_index = header_map[col_name]
            a1_range = gspread.utils.rowcol_to_a1(i, col_index + 1)
            updates.append({"range": f"{sheet_name}!{a1_range}", "values": [[value]]})

    if updates:
        sheet.values().batchUpdate(spreadsheetId=sheet_id, body={"valueInputOption": "USER_ENTERED", "data": updates}).execute()
        print(f"✅ Appended {len(updates)} cells to the Google Sheet.")
    else:
        print("⚠️ No updates applied.")

# ---- Master Runner ----
base_dir = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Mcklenburg\Scraped File"
downloaded_folders = [f for f in os.listdir(base_dir) if f.startswith("Downloaded PDFs")]
downloaded_folders = sorted(downloaded_folders, key=lambda x: os.path.getmtime(os.path.join(base_dir, x)), reverse=True)

latest_folder = os.path.join(base_dir, downloaded_folders[0])
final_pdf_folder = next((os.path.join(latest_folder, f) for f in os.listdir(latest_folder) if "Final PDF" in f), None)
if not final_pdf_folder:
    raise FileNotFoundError("❌ Final PDF folder not found.")

ocr_logs_folder = os.path.join(final_pdf_folder, "OCR Logs")
print(f"📁 Generating OCR logs from: {final_pdf_folder}")
generate_ocr_logs_from_pdfs(final_pdf_folder, ocr_logs_folder)

update_sheet_with_addresses(
    sheet_id="1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ",
    sheet_name="Mecklenburg County",
    credentials_file="credentials.json",
    ocr_folder=ocr_logs_folder
)
