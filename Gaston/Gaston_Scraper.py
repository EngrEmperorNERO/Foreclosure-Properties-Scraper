import os
import time
import zipfile
import glob
import shutil
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import pandas as pd
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
import time
from openpyxl import Workbook

start_time = time.time()

# Paths
firefox_binary_path = r'C:\Program Files\Mozilla Firefox\firefox.exe'
geckodriver_path = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\geckodriver.exe'
download_dir = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Gaston\Scraped File'
options = Options()
options.binary_location = firefox_binary_path

# Set preferences directly on the Options object
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.dir", download_dir)
options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")
options.set_preference("pdfjs.disabled", True)
options.set_preference("browser.download.manager.showWhenStarting", False)
options.set_preference("browser.download.panel.shown", False)
# Set up the WebDriver
service = Service(geckodriver_path)
driver = webdriver.Firefox(service=service, options=options)

# Open the website
url = "https://deeds.gastongov.com/external/LandRecords/protected/v4/SrchNameAdvanced.aspx"
driver.get(url)
driver.maximize_window()

time.sleep(10)

#Click on the acknowledge button if it appears
acknowledge_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "ctl00_btnEmergencyMessagesClose")))
acknowledge_button.click()
time.sleep(2)

# Click on the "Advanced Search" button
advance_search_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "ctl00_NavMenuIdxRec_btnNav_IdxRec_AdvancedName_NEW")))
advance_search_button.click()
time.sleep(2)

# Use yesterday's date as 'Date From' and today's date as 'Date To'
#yesterday = '11/07/2025'
#today = '05/30/2025'

now = datetime.now()
today = now.strftime('%dd/%mm/%YY')
yesterday = (now - timedelta(days=5)).strftime('%dd/%mm/%Yy')


#date from ctl00_cphMain_tcMain_tpNewSearch_ucSrchAdvName_ceFiledFrom_today
#date thru ctl00_cphMain_tcMain_tpNewSearch_ucSrchAdvName_ceFiledThru_today

#Fill Date From Input
date_from_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "ctl00_cphMain_tcMain_tpNewSearch_ucSrchAdvName_txtFiledFrom")))
date_from_input.click()
date_from_input.clear()
date_from_input.send_keys(today)
time.sleep(1)

date_thru_input = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "ctl00_cphMain_tcMain_tpNewSearch_ucSrchAdvName_txtFiledThru")))
date_thru_input.click()
date_thru_input.clear()
date_thru_input.send_keys(today)
time.sleep(1)

kinds_table = driver.find_element(By.ID, "ctl00_cphMain_tcMain_tpNewSearch_ucSrchAdvName_pnlKinds")
# Scroll to kinds_table to make sure elements are in view
driver.execute_script("arguments[0].scrollIntoView(true);", kinds_table)
time.sleep(1)

# Check the 'S/TR' checkbox
str_checkbox = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "ctl00_cphMain_tcMain_tpNewSearch_ucSrchAdvName_repKinds_ctl202_cbKind"))
)
if not str_checkbox.is_selected():
    str_checkbox.click()
time.sleep(1)

#Click Search all matches
search_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "ctl00_cphMain_tcMain_tpNewSearch_ucSrchAdvName_btnInstruments")))
search_button.click()
time.sleep(2)

# Click Date Filled - Sort twice
date_filled_asc_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/form/div[4]/div[2]/div/div[3]/div/div/div[2]/div[3]/div[3]/table/tbody/tr[1]/th[2]/a"))
)
date_filled_asc_button.click()
time.sleep(10)

# Click again to sort in the other direction
date_filled_asc_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/form/div[4]/div[2]/div/div[3]/div/div/div[2]/div[3]/div[3]/table/tbody/tr[1]/th[2]/a"))
)
date_filled_asc_button.click()
time.sleep(10)

from openpyxl import Workbook

# Count the number of clickable Book/Page links
book_page_links = driver.find_elements(By.XPATH, "//a[contains(@id, '_lbDocument_BookPageSuffix')]")
print(f"Total Book/Page links found: {len(book_page_links)}")

results = []

for i in range(2, 2 + len(book_page_links)):  # Starting from ctl02
    try:
        link_id = f"ctl00_cphMain_tcMain_tpInstruments_ucInstrumentsGridV2_cpgvInstruments_ctl{str(i).zfill(2)}_lbDocument_BookPageSuffix"
        print(f"Processing row ID: {link_id}")

        book_page_link = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, link_id))
        )
        book_page_text = book_page_link.text.strip()
        book_page_link.click()
        time.sleep(2)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_cphMain_ctrlResults_cgvResults"))
        )

        grantor_elements = driver.find_elements(By.XPATH, "//table[@id='ctl00_cphMain_ctrlResults_cgvResults']//td[@colspan='2']")
        grantors = ", ".join([el.text.strip() for el in grantor_elements if el.text.strip()])

        returned_to_el = driver.find_element(By.XPATH, "//div[@id='ctl00_cphMain_divReturnToInformation']//table//td")
        returned_to = returned_to_el.text.strip()

        # Get Book/Page from the first row in the References table
        try:
            # Look for the Book/Page anchor that contains a forward slash
            book_page_detail = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((
                    By.XPATH,
                    "/html/body/form/div[4]/div[2]/div/div[3]/div/div[4]/div/div/div/div/table/tbody/tr/td[11]/a"
                ))
            )
            book_page_val = book_page_detail.text.strip()
        except Exception as e:
            print(f"Book/Page scrape failed: {e}")
            book_page_val = "N/A"


        results.append({
            "Book/Page (STR)": book_page_text,
            "Grantors": grantors,
            "Returned To": returned_to,
            "Book/Page (Deed of Trust)": book_page_val
        })

        driver.back()
        time.sleep(2)

        # Refresh link count after back
        book_page_links = driver.find_elements(By.XPATH, "//a[contains(@id, '_lbDocument_BookPageSuffix')]")

    except Exception as e:
        print(f"[Row {i}] Skipped due to error: {e}")
        driver.back()
        time.sleep(2)
        book_page_links = driver.find_elements(By.XPATH, "//a[contains(@id, '_lbDocument_BookPageSuffix')]")
        continue

# ✅ Save to Excel
scrape_date = datetime.now().strftime('%Y-%m-%d')
for r in results:
    r["Date Scraped"] = scrape_date
df = pd.DataFrame(results)

# Move 'Date Scraped' to the front
columns = ["Date Scraped"] + [col for col in df.columns if col != "Date Scraped"]
df = df[columns]

# Split 'Book/Page (Deed of Trust)' into 'Book' and 'Page'
df[['Book', 'Page']] = df['Book/Page (Deed of Trust)'].str.split('/', n=1, expand=True)

# Trim whitespace from Book and Page columns
df['Book'] = df['Book'].str.strip()
df['Page'] = df['Page'].str.strip()

# Drop the original combined column
df.drop(columns=['Book/Page (Deed of Trust)'], inplace=True)

output_path = os.path.join(download_dir, "Gaston_Grantors.xlsx")
df.to_excel(output_path, index=False)
print(f"\n✅ Data saved to: {output_path}")

# Google Sheets append with deduplication
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

# Open the correct sheet and worksheet
spreadsheet = client.open_by_key("1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ")
worksheet = spreadsheet.worksheet("Gaston County")

# Fetch existing values in column A (Book/Page STR) to deduplicate
existing_values = worksheet.col_values(1)[1:]  # skip header

# Filter new records only
new_rows = df[~df["Book/Page (STR)"].isin(existing_values)]

if not new_rows.empty:
    data_to_append = new_rows.values.tolist()
    worksheet.append_rows(data_to_append, value_input_option='RAW')
    print(f"✅ {len(data_to_append)} new rows appended to Google Sheets.")
else:
    print("ℹ️ No new records to append.")

driver.quit()

import subprocess
subprocess.run(["python", r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Gaston\Gaston_Property_Address.py"])
subprocess.run(["python", r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Gaston\Gaston_Property_Address_Parser.py"])