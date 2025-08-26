import os
import time
import random
import gspread
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

# --- Google Sheets Setup ---
SHEET_ID = "1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ"
SHEET_NAME = "Cleveland County"
CREDENTIALS_PATH = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cleveland\credentials.json"

# --- Base Directory Setup ---
base_dir = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cleveland\Scraped and Downloads'
today_str = datetime.now().strftime("%m-%d-%Y")
target_folder = f"Cleveland Scraped File {today_str}"
download_dir = os.path.join(base_dir, target_folder, "Deed of Trust Files")
os.makedirs(download_dir, exist_ok=True)
print(f"📂 Download folder ready: {download_dir}")

# --- Resume Logic: Get only rows with missing PDFs ---
def get_pending_book_page_pairs(sheet_id, sheet_name, credentials_path, download_dir):
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(credentials_path, scopes=scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id).worksheet(sheet_name)

    all_rows = sheet.get_all_records()
    downloaded_files = set(f.lower() for f in os.listdir(download_dir) if f.lower().endswith(".pdf"))

    pending = []
    for row in all_rows:
        book = str(row.get("Book", "")).strip()
        page = str(row.get("Page", "")).strip()
        dt_file = row.get("D/T PDF File", "").strip()

        # Only include if Book/Page present and D/T column is blank
        if book and page and not dt_file:
            expected_filename = f"REAL_PROPERTY_Bk{int(book)}_Pg{str(page).zfill(3)}.pdf"
            if expected_filename.lower() not in downloaded_files:
                pending.append((book, page, expected_filename))

    print(f"🔄 Found {len(pending)} pending Book/Page entries (with empty D/T PDF File column).")
    return pending

# --- Google Sheets Updater ---
def update_dt_pdf_column(book, page, filename, credentials_path, sheet_id, sheet_name):
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_file(credentials_path, scopes=scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(sheet_id).worksheet(sheet_name)

        data = sheet.get_all_values()
        headers = [h.strip().lower() for h in data[0]]
        book_col = headers.index("book")
        page_col = headers.index("page")
        dt_pdf_col = headers.index("d/t pdf file")

        for i, row in enumerate(data[1:], start=2):
            if (
                str(row[book_col]).strip() == str(book)
                and str(row[page_col]).strip() == str(page)
            ):
                sheet.update_cell(i, dt_pdf_col + 1, filename)
                print(f"📝 Updated 'D/T PDF File' for Book {book}, Page {page} at row {i}")
                break
    except Exception as e:
        print(f"⚠️ Failed to update D/T PDF File for Book {book}, Page {page}: {e}")

# --- Selenium Setup ---
firefox_binary_path = r"C:\Program Files\Mozilla Firefox\firefox.exe"
geckodriver_path = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\geckodriver.exe"

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
wait = WebDriverWait(driver, 10)

# --- Navigate to Cleveland Site ---
driver.get("https://us5.courthousecomputersystems.com/ClevelandNCNW/application.asp?resize=true")
driver.maximize_window()
time.sleep(random.randint(5, 10))

# --- Wait for main search frame ---
wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "tabframe0")))

# --- Load rows to process ---
pending_rows = get_pending_book_page_pairs(SHEET_ID, SHEET_NAME, CREDENTIALS_PATH, download_dir)

for i, (book, page, expected_filename) in enumerate(pending_rows, start=1):
    try:
        print(f"[{i}] 🔍 Searching for Book {book}, Page {page} for {expected_filename}")

        book_input = wait.until(EC.presence_of_element_located((By.ID, "booknumber")))
        book_input.clear()
        book_input.send_keys(book)

        page_input = wait.until(EC.presence_of_element_located((By.ID, "pagenumber")))
        page_input.clear()
        page_input.send_keys(page)

        search_button = wait.until(EC.element_to_be_clickable((By.ID, 'search')))
        driver.execute_script("arguments[0].click();", search_button)
        time.sleep(random.randint(5, 10))

        table = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "results")))
        rows = table.find_elements(By.TAG_NAME, "tr")[1:]

        dt_found = False
        for j in range(1, len(rows) + 1):
            try:
                table = driver.find_element(By.CLASS_NAME, "results")
                rows = table.find_elements(By.TAG_NAME, "tr")[1:]
                row = rows[j - 1]

                kind_cell = row.find_element(By.CSS_SELECTOR, "td.col.c8 div")
                kind_value = kind_cell.get_attribute("title").strip()

                if "D/T" in kind_value:
                    img_icon = row.find_element(By.CSS_SELECTOR, "td.col.c2 img[title='Document image is available']")
                    driver.execute_script("arguments[0].scrollIntoView(true);", img_icon)
                    time.sleep(random.randint(5, 10))
                    img_icon.click()
                    print(f"[{i}] ✅ Clicked image for row {j} (Kind: {kind_value})")
                    time.sleep(random.randint(5, 10))

                    driver.switch_to.default_content()
                    outer_iframe = wait.until(EC.presence_of_element_located((By.NAME, "tabframe1")))
                    driver.switch_to.frame(outer_iframe)

                    inner_switched = False
                    for attempt in range(5):
                        try:
                            inner_iframe = WebDriverWait(driver, 3).until(
                                EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'viewimageframe.asp')]"))
                            )
                            driver.switch_to.frame(inner_iframe)
                            print(f"[{i}] ✅ Switched to image iframe.")
                            inner_switched = True
                            break
                        except:
                            time.sleep(1)

                    if not inner_switched:
                        raise Exception(f"[{i}] ❌ inner iframe not found after retries.")

                    download_link = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Download Image')]"))
                    )
                    print(f"[{i}] 👀 'Download Image' link found.")

                    if not download_link.is_displayed():
                        driver.execute_script("arguments[0].scrollIntoView(true);", download_link)
                        time.sleep(random.randint(5, 10))

                    driver.execute_script("arguments[0].click();", download_link)
                    print(f"[{i}] 📅 Download triggered for {expected_filename}.")
                    time.sleep(random.randint(5, 10))

                    update_dt_pdf_column(
                        book=book,
                        page=page,
                        filename=expected_filename,
                        credentials_path=CREDENTIALS_PATH,
                        sheet_id=SHEET_ID,
                        sheet_name=SHEET_NAME
                    )

                    driver.switch_to.default_content()
                    tab_link = wait.until(EC.presence_of_element_located((By.ID, "tab0")))
                    driver.execute_script("arguments[0].click();", tab_link)
                    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "tabframe0")))
                    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "results")))
                    print(f"[{i}] ♻ Returned to table view.")
                    dt_found = True
                    break

            except Exception as e:
                print(f"[{i}] ⚠️ Row scan error: {e}")

        if not dt_found:
            print(f"[{i}] ❌ No row with D/T found.")

    except Exception as e:
        print(f"[{i}] ❌ Search failed: {e}")

driver.quit()
