import os
import re
import pytesseract
from PIL import Image
import pdf2image
from datetime import datetime
from PIL import Image

Image.MAX_IMAGE_PIXELS = None

# Tesseract path
pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Tesseract\tesseract.exe"
)

# Cleveland base dir
base_dir = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cleveland\Scraped and Downloads"

# Get the latest folder
folders = [f for f in os.listdir(base_dir) if f.startswith("Cleveland Scraped File ")]
if not folders:
    raise FileNotFoundError("No 'Cleveland Scraped File' folders found.")

latest_folder = max(
    folders,
    key=lambda f: datetime.strptime(f.replace("Cleveland Scraped File ", ""), "%m-%d-%Y")
)
pdf_folder = os.path.join(base_dir, latest_folder)
print(f"📁 Using folder: {pdf_folder}")

# Prepare file list
pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]
ocr_log_dir = os.path.join(pdf_folder, "OCR - STR Book and Page Logs")
os.makedirs(ocr_log_dir, exist_ok=True)

# Regex patterns
recording_pat = re.compile(
    r'(?:record(?:ed)?\s+on\s+)?(?:\w+\s+\d{1,2},\s+\d{4},\s+)?in\s+Book(?:\s+No\.?)?\s+(\d{3,6})[,\s]+at\s+Page[:\s]+(\d{1,6})',
    re.IGNORECASE
)
combined_pat = re.compile(
    r'Book[:\s_]*?(\d{3,6})[,\s]+Page[:\s_]*?(\d{1,6})'
    r'|Book\s+(\d{3,6})\s+Page\s+(\d{1,6})'
    r'Book\s*No\.?\s*(\d{3,6}),?\s*at\s*Page\s*0*(\d{1,6})'
    r'|Book\s*No\.?\s*(\d{3,6}),?\s*at\s*Page\s*0*(\d{1,6})'
    r'|Book(?:\s+Cleveland)?[:\s]+(\d{3,6}).*?Page[:\s]+(\d{1,6})',
    re.IGNORECASE
)
# Add pattern for "Recorded on June 23, 2023, in Book No Cleveland 1901, at Page 491"
cleveland_pat = re.compile(
    r'Recorded on \w+ \d{1,2}, \d{4}, in Book No Cleveland\s*(\d{3,6}), at Page\s*(\d{1,6})',
    re.IGNORECASE
)

# --- OCR processor ---
def generate_ocr_log(pdf_path, output_folder, filename_base):
    try:
        images = pdf2image.convert_from_path(pdf_path, dpi=300)
        full_text = ""
        for img in images:
            text = pytesseract.image_to_string(img)
            full_text += text + "\n\n"

        os.makedirs(output_folder, exist_ok=True)
        output_path = os.path.join(output_folder, f"{filename_base}.txt")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(full_text.strip())
        return full_text
    except Exception as e:
        print(f"⚠️ Error generating OCR log for {filename_base}: {e}")
        return ""

# --- Process PDFs ---
results = []
missing_book_page = []

for pdf_file in pdf_files:
    filename_base = os.path.splitext(pdf_file)[0]
    pdf_path = os.path.join(pdf_folder, pdf_file)
    txt_path = os.path.join(ocr_log_dir, f"{filename_base}.txt")

    ocr_text = generate_ocr_log(pdf_path, ocr_log_dir, filename_base)
    if not ocr_text:
        print(f"❌ Skipping due to OCR failure: {pdf_file}")
        continue

    # Strip header lines to avoid misfires
    lines = ocr_text.splitlines()
    start_idx = 0
    for i, line in enumerate(lines):
        if "substitute trustee" in line.lower():
            start_idx = i
            break
    main_text = " ".join(lines[start_idx:]).replace("\n", " ").strip()

    # Extract Book/Page
    book = page = None
    m1 = recording_pat.search(main_text)
    if m1:
        book, page = m1.group(1), m1.group(2)
    else:
        m2 = combined_pat.search(main_text)
        if m2:
            groups = m2.groups()
            for j in range(0, len(groups), 2):
                if groups[j] and groups[j + 1]:
                    book, page = groups[j], groups[j + 1]
                    break

    if book or page:
        print(f"{pdf_file} -> 📘 Book: {book}, 📄 Page: {page}")
        results.append({"file": pdf_file, "book": book, "page": page})
    else:
        print(f"⚠️ Could not extract Book/Page from {pdf_file}")
        missing_book_page.append({
            "file": pdf_file,
            "date_scraped": datetime.today().strftime("%Y-%m-%d")
        })
import gspread
from google.oauth2.service_account import Credentials

import gspread
from google.oauth2.service_account import Credentials

def update_sheet_book_page_by_filename(results, credentials_path, sheet_id, sheet_name):
    if not results:
        print("⚠️ No results to update.")
        return

    # Authenticate
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(credentials_path, scopes=scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id).worksheet(sheet_name)

    # Get all existing values
    all_data = sheet.get_all_values()
    headers = all_data[0]
    rows = all_data[1:]

    # Determine column indices
    filename_col = headers.index("STR PDF")
    book_col = headers.index("Book")
    page_col = headers.index("Page")

    # Create row map from PDF filename
    row_lookup = {row[filename_col].strip(): idx + 2 for idx, row in enumerate(rows)}  # +2 for 1-based + header

    update_cells = []
    for entry in results:
        filename = entry["file"].strip()
        row_num = row_lookup.get(filename)
        if row_num:
            sheet.update_cell(row_num, book_col + 1, entry["book"])  # gspread is 1-based
            sheet.update_cell(row_num, page_col + 1, entry["page"])
            print(f"✅ Updated row {row_num} with Book: {entry['book']}, Page: {entry['page']}")
        else:
            print(f"⚠️ No matching row found for {filename}")

    print("✅ Finished updating Book and Page columns.")

# --- Update Google Sheet ---
update_sheet_book_page_by_filename(
    results=results,
    credentials_path=r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cleveland\credentials.json",
    sheet_id="1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ",
    sheet_name="Cleveland County"
)

