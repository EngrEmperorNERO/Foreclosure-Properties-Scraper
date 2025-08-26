import os
import time
import re
from datetime import datetime, timedelta
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- Paths ---
firefox_binary_path = r'C:\Program Files\Mozilla Firefox\firefox.exe'
geckodriver_path = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\geckodriver.exe'

# --- Set dynamic dated download folder and clean up old PDFs ---
base_dir = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cabarrus'
today_str = datetime.now().strftime('%m-%d-%Y')
download_dir = os.path.join(base_dir, f"Cabarrus Scraped File {today_str}")
os.makedirs(download_dir, exist_ok=True)



# Set Excel output path
output_excel = os.path.join(download_dir, "cabarrus_subt_all.xlsx")

# --- Selenium Setup ---
options = Options()
options.binary_location = firefox_binary_path
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.dir", download_dir)
options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")
options.set_preference("pdfjs.disabled", True)
options.set_preference("browser.download.manager.showWhenStarting", False)
options.set_preference("browser.download.panel.shown", False)

service = Service(geckodriver_path)
driver = webdriver.Firefox(service=service, options=options)

# --- Start ---
driver.get("https://www.cabarrusncrod.org/")
driver.maximize_window()
time.sleep(2)
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(2)

# Accept terms
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[2]/div/div[6]/a"))
).click()
time.sleep(2)

# Click Full System
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div/div[5]/a"))
).click()
time.sleep(5)

# Search by Recorded Date
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "TAB_42"))
).click()
time.sleep(2)

# Set Dates
now = datetime.now()
today = now.strftime('%m/%d/%Y')
start_date = (now - timedelta(days=4)).strftime('%m/%d/%Y')

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'TRG_98'))).send_keys(today)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'TRG_99'))).send_keys(today)

# Instrument Type
instrument_input = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'TRG_95')))
instrument_input.clear()
instrument_input.send_keys("SUB-T")
driver.find_element(By.TAG_NAME, 'body').click()
time.sleep(1)

# Search
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'VWG_25'))).click()
time.sleep(5)

# --- Get all rows ---
all_data = []
row_count = len(driver.find_elements(By.XPATH, "//div[starts-with(@id, 'VWGROW2_152_R')]"))
print(f"🔎 Found {row_count} rows to process...")

for i in range(row_count):
    try:
        # Re-fetch rows fresh each time
        rows = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//div[starts-with(@id, 'VWGROW2_152_R')]"))
        )
        row = rows[i]
        cols = row.find_elements(By.XPATH, "./div[starts-with(@id, 'VWG_152_D')]")

        if len(cols) < 8:
            print(f"⚠️ Skipping row {i+1}: incomplete data")
            continue

        rec_date = cols[2].text.strip()
        instr    = cols[3].text.strip()
        book     = cols[4].text.strip()
        page     = cols[5].text.strip()
        type_    = cols[6].text.strip()
        desc     = cols[7].text.strip()

        print(f"\n--- Row {i+1} --- Instr #: {instr} | Book: {book} | Page: {page}")

        # Click to open image view
        driver.execute_script("arguments[0].scrollIntoView(true);", cols[2])
        original_window = driver.current_window_handle
        cols[2].click()
        time.sleep(10)

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "VWG_185")))

        # Confirm if image button is present
        try:
            image_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "VWG_159")))
        except:
            print(f"⚠️ Image button not found or not clickable for row {i+1}, skipping...")
            continue

        if not image_button.is_displayed():
            print(f"⚠️ Image button not visible for row {i+1}, skipping...")
            continue

        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", image_button)

        # Wait and switch to new tab
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
        new_window = [w for w in driver.window_handles if w != original_window][0]
        driver.switch_to.window(new_window)
        time.sleep(2)

        # Frame and grantor/grantee parsing
        try:
            driver.switch_to.frame(1)
            detail_div = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, f"I{instr}"))  # Dynamic ID
            )
            raw_html = detail_div.get_attribute("innerHTML")
            soup = BeautifulSoup(raw_html, "html.parser")

            grantor = []
            grantee = []
            current_section = None

            for tag in soup.children:
                if tag.name == "strong":
                    if "Grantor" in tag.text:
                        current_section = "grantor"
                    elif "Grantee" in tag.text:
                        current_section = "grantee"
                elif tag.name == "br":
                    continue
                elif isinstance(tag, str):
                    text = tag.strip()
                    if text:
                        if current_section == "grantor":
                            grantor.append(text)
                        elif current_section == "grantee":
                            grantee.append(text)

            grantor_text = "; ".join(grantor) if grantor else "N/A"
            grantee_text = "; ".join(grantee) if grantee else "N/A"
        except Exception as e:
            print(f"⚠️ Frame parsing failed on row {i+1}: {e}")
            grantor_text = "N/A"
            grantee_text = "N/A"

        all_data.append({
            "Rec Date": rec_date,
            "Instr #": instr,
            "Book": book,
            "Page": page,
            "Type": type_,
            "Description": desc,
            "Grantor(s)": grantor_text,
            "Grantee(s)": grantee_text
        })

        # Close popup tab and return
        driver.close()
        driver.switch_to.window(original_window)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "VWG_25")))
        time.sleep(2)

        # ✅ Confirm row completion
        print(f"✅ Finished processing row {i+1}")

    except Exception as e:
        print(f"❌ Error on row {i+1}: {e}")
        try:
            for handle in driver.window_handles:
                if handle != driver.current_window_handle:
                    driver.switch_to.window(handle)
                    driver.close()
            driver.switch_to.window(original_window)
        except:
            pass
        continue

# --- Delete all PDFs in the folder before starting ---
deleted_files = 0
for f in os.listdir(download_dir):
    file_path = os.path.join(download_dir, f)
    if f.lower().endswith(".pdf") and os.path.isfile(file_path):
        try:
            os.remove(file_path)
            deleted_files += 1
        except Exception as e:
            print(f"⚠️ Could not delete {f}: {e}")

print(f"🧹 Cleared {deleted_files} PDF(s) in: {download_dir}")

# --- Create DataFrame and enrich data ---
scrape_date = datetime.now().strftime('%m/%d/%Y')
df = pd.DataFrame(all_data)

# Insert Date Scraped as first column
df.insert(0, "Date Scraped", scrape_date)

# Extract Book and Page number from the 'Description' column
def extract_book_page(desc):
    match = re.search(r'BOOK\s+(\d+)\s+PG\s+(\d+)', desc.upper())
    if match:
        return match.group(1), match.group(2)
    return "", ""

def extract_and_format_book_page(desc):
    match = re.search(r'BOOK\s+(\d+)\s+PG\s+(\d+)', desc.upper())
    if match:
        book = match.group(1).zfill(5)
        page = match.group(2).zfill(4)
        return book, page
    return "", ""

df["Book Number (D/T)"], df["Page Number (D/T)"] = zip(*df["Description"].map(extract_and_format_book_page))

# --- Save to Excel ---
df.to_excel(output_excel, index=False)
print(f"\n📁 All data saved to: {output_excel}")

# --- Upload to Google Sheets ---
import gspread
from google.oauth2.service_account import Credentials

SHEET_ID = "1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ"
SHEET_NAME = "Cabarrus County"
CREDENTIALS_FILE = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cabarrus\credentials.json"

scope = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)

# Convert all to strings and pad where needed
values = df.fillna("").astype(str)

# Add single quotes only to Book and Page numbers to preserve leading zeros
def force_text_format(row):
    row = list(row)  # ensure mutable
    book = row[values.columns.get_loc("Book Number (D/T)")]
    page = row[values.columns.get_loc("Page Number (D/T)")]
    row[values.columns.get_loc("Book Number (D/T)")] = "'" + book
    row[values.columns.get_loc("Page Number (D/T)")] = "'" + page
    return row

formatted_values = [force_text_format(row) for row in values.values.tolist()]
sheet.append_rows(formatted_values, value_input_option="USER_ENTERED")

# --- Close browser ---
driver.quit()
print("🛑 Browser session closed.")

import subprocess

subprocess.run(["python", r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cabarrus\Cabarrus_DT_Download.py"])
subprocess.run(["python", r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cabarrus\Cabarrus_DT_Parsing.py"])
