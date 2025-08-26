import os
import glob
import time
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Setup
firefox_binary_path = r'C:\Program Files\Mozilla Firefox\firefox.exe'
geckodriver_path = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\geckodriver.exe'
download_dir = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Gaston\Scraped File'
options = Options()
options.binary_location = firefox_binary_path
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.dir", download_dir)
options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")
options.set_preference("pdfjs.disabled", True)
options.set_preference("browser.download.manager.showWhenStarting", False)

service = Service(geckodriver_path)
driver = webdriver.Firefox(service=service, options=options)

# Load website
base_url = "https://deeds.gastongov.com/external/LandRecords/protected/v4/SrchBookPage.aspx"
driver.get(base_url)
driver.maximize_window()
time.sleep(2)

WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "ctl00_btnEmergencyMessagesClose"))).click()
time.sleep(2)

# Google Sheets setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open_by_key("1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ").worksheet("Gaston County")
df = pd.DataFrame(sheet.get_all_records())

# Loop through rows
for index, row in df.iterrows():
    # Skip if PDF was already downloaded
    if row.get('Downloaded PDF'):
        print(f"⏭️ Skipping Row {index + 1} — Already downloaded: {row['Downloaded PDF']}")
        continue

    book = str(row['Book'])
    page = str(row['Page'])
    date_scraped = row.get('Date Scraped', '')
    date_scraped_str = pd.to_datetime(date_scraped).strftime('%Y-%m-%d') if date_scraped else 'UnknownDate'
    target_dir = os.path.join(download_dir, date_scraped_str)
    os.makedirs(target_dir, exist_ok=True)

    print(f"\n📘 Processing Row {index + 1} — Book {book}, Page {page}, Date Scraped: {date_scraped_str}")
    
    try:
        # Always reload Book/Page search tab
        driver.get(base_url)
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ctl00_NavMenuIdxRec_btnNav_IdxRec_BookPage_NEW"))
        ).click()
        time.sleep(2)

        # Freshly locate fields each time
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_cphMain_tcMain_tpNewSearch_ucSrchBkPg_txtBookNumber"))).clear()
        driver.find_element(By.ID, "ctl00_cphMain_tcMain_tpNewSearch_ucSrchBkPg_txtBookNumber").send_keys(book)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_cphMain_tcMain_tpNewSearch_ucSrchBkPg_txtPageNumber"))).clear()
        driver.find_element(By.ID, "ctl00_cphMain_tcMain_tpNewSearch_ucSrchBkPg_txtPageNumber").send_keys(page)

        driver.find_element(By.ID, "ctl00_cphMain_tcMain_tpNewSearch_ucSrchBkPg_btnSearch").click()

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,
            "//table[@id='ctl00_cphMain_tcMain_tpInstruments_ucInstrumentsGridV2_cpgvInstruments']//tr[contains(@class, 'cottPagedGridViewRowStyle')]"
        )))

        rows = driver.find_elements(By.XPATH, "//tr[contains(@class, 'cottPagedGridViewRowStyle')]")

        for r in rows:
            cells = r.find_elements(By.TAG_NAME, "td")
            if len(cells) < 10:
                continue

            kind = cells[3].text.strip().upper()
            if kind != "D/T":
                continue

            print(f"✅ Found D/T row: Book {book}, Page {page}")

            try:
                r.find_element(By.XPATH, ".//input[contains(@src, 'document.png')]").click()
                time.sleep(2)

                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                WebDriverWait(driver, 20).until(EC.invisibility_of_element((By.ID, "imageLoadOverlay")))
                time.sleep(10)

                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and @value='Save Document as PDF']"))
                ).click()
                print("✅ Clicked 'Save Document as PDF' button.")

                time.sleep(30)

                print("⏳ Waiting for PDF to download...")
                timeout = 30
                poll = 0
                pdf_file = None

                while poll < timeout:
                    files = glob.glob(os.path.join(download_dir, "*.pdf"))
                    part_files = glob.glob(os.path.join(download_dir, "*.part"))
                    if files and not part_files:
                        pdf_file = max(files, key=os.path.getctime)
                        break
                    time.sleep(1)
                    poll += 1

                if not pdf_file:
                    print("❌ PDF download did not complete.")
                else:
                    new_name = f"Book_{book}_Page_{page}.pdf"
                    new_path = os.path.join(target_dir, new_name)

                    if os.path.exists(new_path):
                        os.remove(new_path)
                    os.rename(pdf_file, new_path)
                    print(f"📁 Saved: {new_path}")

                    # ✅ Update Google Sheet with downloaded filename
                    sheet.update_cell(index + 2, 7, new_name)  # index + 2 because sheet is 1-indexed and row 1 is the header
                    print("📌 Updated sheet with downloaded filename.")
                    break
            except Exception as e:
                print(f"⚠️ Error during PDF download: {e}")
                break

    except Exception as e:
        print(f"❌ Error processing Book {book}, Page {page}: {e}")

driver.quit()
