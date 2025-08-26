import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import gspread
from google.oauth2.service_account import Credentials
import unicodedata
from PyPDF2 import PdfReader
import re

# --- Paths ---
firefox_binary_path = r'C:\Program Files\Mozilla Firefox\firefox.exe'
geckodriver_path = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\geckodriver.exe'

# --- Set dynamic dated download folder ---
base_dir = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cabarrus'


def get_latest_scraped_folder(base_path):
    folder_pattern = re.compile(r"Cabarrus Scraped File (\d{2}-\d{2}-\d{4})")
    dated_folders = []

    for name in os.listdir(base_path):
        match = folder_pattern.match(name)
        if match:
            try:
                date_obj = datetime.strptime(match.group(1), "%m-%d-%Y")
                dated_folders.append((date_obj, os.path.join(base_path, name)))
            except ValueError:
                continue

    if not dated_folders:
        raise Exception("❌ No valid dated folders found in base directory.")

    latest_folder = max(dated_folders, key=lambda x: x[0])[1]
    print(f"📁 Using latest scraped folder: {os.path.basename(latest_folder)}")
    return latest_folder

# Use latest scraped folder as download directory
download_dir = get_latest_scraped_folder(base_dir)


# --- Selenium Setup ---
options = Options()
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.dir", download_dir)
options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")
options.set_preference("pdfjs.disabled", True)
options.set_preference("browser.download.useDownloadDir", True)
options.set_preference("browser.download.panel.shown", False)
options.set_preference("browser.download.manager.showWhenStarting", False)
options.set_preference("browser.download.always_ask_before_handling_new_types", False)

service = Service(geckodriver_path)
driver = webdriver.Firefox(service=service, options=options)


def normalize_text(value: str) -> str:
    if not isinstance(value, str):
        value = str(value)
    return unicodedata.normalize('NFKD', value).strip().replace('\u200b', '').replace('\xa0', '')


# --- Google Sheets config ---
SHEET_ID = "1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ"
SHEET_NAME = "Cabarrus County"
CREDENTIALS_FILE = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cabarrus\credentials.json"

# --- Connect to Google Sheets ---
scope = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)

# --- Fetch all valid Book/Page entries from Google Sheet ---
records = sheet.get_all_values()  # Raw strings including header
header = records[0]
book_col = header.index("Book Number (D/T)")
page_col = header.index("Page Number (D/T)")
deed_col_index = header.index("Deed of Trust PDF") + 1  # Column L (1-based index)

book_page_keys = []

for i, row in enumerate(records[1:], start=2):  # start=2 to align with actual sheet row index
    if len(row) > max(book_col, page_col, deed_col_index - 1):
        book = row[book_col].strip()
        page = row[page_col].strip()
        deed_filename = row[deed_col_index - 1].strip()
        if book and page and not deed_filename:
            book_page_keys.append((book, page, i))  # 'i' is the real sheet row index

if not book_page_keys:
    print("❌ No valid Book/Page entries found in Google Sheet.")
    driver.quit()
    exit()

print(f"✅ Loaded {len(book_page_keys)} Book/Page entries from Google Sheet.")

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

# --- Loop through each Book/Page pair ---
for index, (book, page, sheet_row_index) in enumerate(book_page_keys, 1):
    print(f"\n🔍 [{index}] Searching Book: {book}, Page: {page}")

    # Step 1: Click Book/Page tab (TAB_40)
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "TAB_40"))
    ).click()
    time.sleep(2)

    # Step 2: Fill Book number
    book_input = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "TRG_81"))
    )
    book_input.clear()
    book_input.send_keys(book)

    # Step 3: Fill Page number
    page_input = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "TRG_80"))
    )
    page_input.clear()
    page_input.send_keys(page)

    # Remove focus from input to trigger update
    driver.find_element(By.TAG_NAME, 'body').click()
    time.sleep(5)

    # Step 4: Click Search
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "VWG_25"))
    ).click()

    print("📄 Search submitted...")
    time.sleep(5)



    # Step 5: Click checkboxes for all matching Book/Page rows
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "VWGDGVBODY_152"))
        )
        print("✅ Table found. Scanning rows...")

        matched = False
        row_index = 0

        while True:
            try:
                rows = driver.find_elements(By.XPATH, "//div[starts-with(@id, 'VWGROW2_152_R')]")
                if row_index >= len(rows):
                    break

                row = rows[row_index]
                cols = row.find_elements(By.XPATH, "./div[starts-with(@id, 'VWG_152_D')]")
                if len(cols) < 4:
                    row_index += 1
                    continue

                # Extract Book and Page values
                try:
                    row_book = cols[2].find_element(By.TAG_NAME, "span").text.strip()
                except:
                    row_book = cols[2].text.strip()

                try:
                    row_page = cols[3].find_element(By.TAG_NAME, "span").text.strip()
                except:
                    row_page = cols[3].text.strip()

                # Retry if empty
                if not row_book or not row_page:
                    time.sleep(2)
                    try:
                        row_book = cols[2].find_element(By.TAG_NAME, "span").text.strip()
                    except:
                        row_book = cols[2].text.strip()
                    try:
                        row_page = cols[3].find_element(By.TAG_NAME, "span").text.strip()
                    except:
                        row_page = cols[3].text.strip()

                print(f"🔎 Row {row_index + 1}: Book={row_book}, Page={row_page}")

                time.sleep(2)
                
                def wait_for_file_complete(path, timeout=20):
                    prev_size = -1
                    start_time = time.time()
                    while time.time() - start_time < timeout:
                        if os.path.exists(path):
                            curr_size = os.path.getsize(path)
                            if curr_size == prev_size and curr_size > 0:
                                return True
                            prev_size = curr_size
                        time.sleep(0.5)
                    return False

                if row_book == book and row_page == page:
                    checkbox_cell = cols[2]
                    driver.execute_script("arguments[0].click();", checkbox_cell)
                    print(f"✅ Clicked checkbox on Row {row_index + 1}")
                    matched = True

                    try:
                        # --- Click image button to trigger download ---
                        time.sleep(2)
                        image_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "VWG_159"))
                        )
                        driver.execute_script("arguments[0].click();", image_button)
                        print("📥 Clicked image button (VWG_159)")

                        # --- Switch to popup ---
                        # --- Switch to popup ---
                        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
                        main_window = driver.current_window_handle
                        popup_window = [w for w in driver.window_handles if w != main_window][0]
                        driver.switch_to.window(popup_window)
                        print("🪟 Switched to popup window")

                        # --- Track files before download ---
                        existing_files = set(os.listdir(download_dir))

                        # --- Wait for new file to download ---
                        timeout = 15
                        start_time = time.time()
                        downloaded_file = None

                        while time.time() - start_time < timeout:
                            current_files = set(os.listdir(download_dir))
                            new_files = current_files - existing_files
                            new_pdfs = [f for f in new_files if f.lower().endswith(".pdf")]
                            if new_pdfs:
                                downloaded_file = new_pdfs[0]
                                break
                            time.sleep(5)

                        # --- Wait for file to finish writing ---
                        if downloaded_file:
                            downloaded_path = os.path.join(download_dir, downloaded_file)
                            if wait_for_file_complete(downloaded_path):
                                new_filename = f"{book}_{page}.pdf"
                                dst_path = os.path.join(download_dir, new_filename)
                                os.rename(downloaded_path, dst_path)
                                print(f"✅ Download complete and saved as: {new_filename}")
                                sheet.update_cell(sheet_row_index, deed_col_index, new_filename)
                            else:
                                print(f"❌ File '{downloaded_file}' was not fully written before timeout.")
                        else:
                            print("❌ No new PDF file found after clicking image button.")
                            sheet.update_cell(sheet_row_index, deed_col_index, "Deed of Trust not Found")

                        # --- Close popup and return ---
                        driver.close()
                        driver.switch_to.window(main_window)
                        print("↩️ Closed popup and returned to main window")

                    except Exception as e:
                        print(f"⚠️ Error during auto-download/rename process: {e}")

                    break  # Done after one match

                row_index += 1

                if row_index % 10 == 0:
                    try:
                        scroll_script = """
                            const el = document.getElementById('VWGDGVBODY_152');
                            if (el) el.scrollTop += 100;
                        """
                        driver.execute_script(scroll_script)
                        print(f"↕️ Scrolled table after Row {row_index}")
                        time.sleep(2)
                    except Exception as scroll_err:
                        print(f"⚠️ Scroll failed at Row {row_index}: {scroll_err}")

            except Exception as e:
                print(f"⚠️ Retrying row {row_index + 1} due to dynamic reload: {e}")
                time.sleep(2)
                continue

        if not matched:
            print(f"❌ No matching Book/Page checkbox clicked for Book={book}, Page={page}")

    except Exception as e:
        print(f"❌ Error accessing table rows or clicking checkbox: {e}")

    time.sleep(5)


    # Step 6: Return to Index Search tab (TAB_11)
    try:
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "TAB_11"))
        ).click()
        print("↩️ Returned to Index Search tab.")
        time.sleep(2)
    except Exception as e:
        print(f"⚠️ Failed to return to Index Search tab: {e}")
        break
driver.quit()

import subprocess
subprocess.run(["python", r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cabarrus\Cabarrus_DT_Parsing.py"])
