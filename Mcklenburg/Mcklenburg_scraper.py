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

start_time = time.time()

# Paths
firefox_binary_path = r'C:\Program Files\Mozilla Firefox\firefox.exe'
geckodriver_path = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\geckodriver.exe'
download_dir = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Mcklenburg\Scraped File'
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
url = "https://meckrod.manatron.com/"
driver.get(url)
driver.maximize_window()
# Wait for the page to load and accept terms
time.sleep(5)
enter_site = driver.find_element(By.XPATH, '//*[@id="cph1_lnkAccept"]')
enter_site.click()
#Real Estate Button
real_estate_button = driver.find_element(By.XPATH, '/html/body/div[3]/form/div[3]/div[1]/div[2]/div/ul/li[13]/a').click()
time.sleep(5)
# Search Real Estate Index
real_estate_index_button = driver.find_element(By.XPATH, '/html/body/div[3]/form/div[3]/div[1]/div[2]/div[2]/div/div/div/ul/li[3]/a').click()
time.sleep(5)
# Use yesterday's date as 'Date From' and today's date as 'Date To'
yesterday = '06/19/2025'
#today = '05/30/2025'

#yesterday = (datetime.now() - timedelta(days=1)).strftime('%m/%d/%Y')
today = datetime.now().strftime('%m/%d/%Y')

# Input 'Date From'
date_selector = driver.find_element(By.XPATH, '/html/body/div[3]/form/div[3]/div[3]/table/tbody/tr[3]/td/table[1]/tbody/tr[9]/td[2]/table/tbody/tr/td[1]/input')
date_selector.clear()
date_selector.send_keys(today)
time.sleep(2)
# Input 'Date To'
date_selector = driver.find_element(By.XPATH, '/html/body/div[3]/form/div[3]/div[3]/table/tbody/tr[3]/td/table[1]/tbody/tr[9]/td[4]/table/tbody/tr/td[1]/input')
date_selector.clear()
date_selector.send_keys(today)
time.sleep(2)
# Scroll to the checkbox frame containing document types
checkbox_container = driver.find_element(By.XPATH, '/html/body/div[3]/form/div[3]/div[3]/table/tbody/tr[3]/td/table[1]/tbody/tr[12]/td[2]/div')
driver.execute_script("arguments[0].scrollIntoView(true);", checkbox_container)
time.sleep(2)


# Find and click the checkbox for "SUBSTITUTION TRUSTEE"
checkbox_label = driver.find_element(By.XPATH, '//label[text()="SUBSTITUTION TRUSTEE"]')
checkbox_id = checkbox_label.get_attribute("for")  # e.g. 'cphNoMargin_f_dclDocType_368'
checkbox_input = driver.find_element(By.ID, checkbox_id)
if not checkbox_input.is_selected():
    checkbox_input.click()
    print("Checkbox for 'SUBSTITUTION TRUSTEE' clicked.")
else:
    print("Checkbox already selected.")
#Click the search button
search = driver.find_element(By.XPATH, '//*[@id="cphNoMargin_SearchButtons2_btnSearch__3"]')
search.click()
time.sleep(5)
# Wait for the table to load
table = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/form/div[3]/div[3]/table/tbody/tr[4]/td/div/div/table/tbody/tr[2]/td/table'))
)

# Check if the table exists; if not, quit and print debug message
if not table:
    print(f"No New List for ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")
    driver.quit()
    exit()
    
# Prepare the data container
data = []
# Get all data rows (skipping the header row)
rows = table.find_elements(By.XPATH, './/tr[position() > 1]')
print(f"Number of rows found: {len(rows)}")
data = []
for row in rows:
    try:
        cols = row.find_elements(By.TAG_NAME, "td")
        # Extract values
        instrument = cols[4].text.strip().replace('\n', ' ')
        date_filed = cols[8].text.strip()
        doc_type = cols[9].text.strip()
        party_names = cols[11].text.strip().replace('\n', ' | ')  # Format nicely
        print(f"Extracted -> Instrument: {instrument}, Date Filed: {date_filed}, Doc Type: {doc_type}, Parties: {party_names}")
        data.append({
            "Instrument # Book-Page": instrument,
            "Date Filed": date_filed,
            "Document Type": doc_type,
            "Party Name(s)": party_names
        })
    except Exception as e:
        print(f"Error extracting row: {e}")


# Convert to DataFrame
df = pd.DataFrame(data)
df = df.fillna("")
# Extract Instrument # and Book - Page correctly
# Format: "2025044282 39568- 294" → Instrument #: 2025044282, Book - Page: 39568 - 294
df[["Instrument #", "Book - Page"]] = df["Instrument # Book-Page"].str.extract(r'^(\d+)\s+(.*)$')

# Drop the original combined column
df.drop(columns=["Instrument # Book-Page"], inplace=True)

# Reorder columns if desired
df = df[["Instrument #", "Book - Page", "Date Filed", "Document Type", "Party Name(s)"]]

# Save to Excel
output_path = os.path.join(download_dir, "instrument_book_page_results.xlsx")
df.to_excel(output_path, index=False)

print(f"Data saved to {output_path}")

# Locate the result table body once
results_table_xpath = '/html/body/div[3]/form/div[3]/div[3]/table/tbody/tr[4]/td/div/div/table/tbody/tr[2]/td/table/tbody[2]'
results_table = driver.find_element(By.XPATH, results_table_xpath)
rows = results_table.find_elements(By.TAG_NAME, 'tr')


def wait_for_download(dir_path, timeout=60):
    print("Waiting for file download to complete...")
    seconds = 0
    while seconds < timeout:
        time.sleep(1)
        if not any(fname.endswith(".part") for fname in os.listdir(dir_path)):
            print("Download finished.")
            return
        seconds += 1
    print("Download timed out.")

# Get the number of rows in the table
initial_rows = driver.find_elements(By.XPATH, f"{results_table_xpath}/tr")
row_index = 0
page_number = 1
last_first_row_text = None

while True:
    print(f"\n📄 Processing Page {page_number}")

    # === Scraping rows on current page ===
    try:
        results_table = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, results_table_xpath))
        )
        rows = results_table.find_elements(By.TAG_NAME, 'tr')

        if not rows:
            print("⚠️ No rows found on this page.")
            break

        # Detect repeated page
        current_first_row_text = rows[0].text.strip()
        if current_first_row_text == last_first_row_text:
            print("🚫 Detected same first row as previous page — likely last page reached. Stopping pagination.")
            break
        last_first_row_text = current_first_row_text

        row_index = 0
        while row_index < len(rows):
            try:
                row = rows[row_index]
                tds = row.find_elements(By.TAG_NAME, "td")
                if not tds:
                    row_index += 1
                    continue
                view_td = next((td for td in tds if 'View' in td.text), None)
                if not view_td:
                    row_index += 1
                    continue
                view_button = view_td.find_element(By.CLASS_NAME, "staticLink")
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", view_button)
                time.sleep(0.5)
                original_window = driver.current_window_handle
                view_button.click()
                WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
                new_window = [w for w in driver.window_handles if w != original_window][0]
                driver.switch_to.window(new_window)

                WebDriverWait(driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it((By.ID, 'ImageViewer1_ifrLTViewer'))
                )
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.ID, 'form1'))
                )
                get_image_button = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.ID, 'btnProcessNow'))
                )
                driver.execute_script("arguments[0].click();", get_image_button)
                time.sleep(30)

                WebDriverWait(driver, 20).until(
                    EC.visibility_of_element_located((By.ID, "dlgImageWindow"))
                )
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.ID, 'ifrImageWindow'))
                )
                driver.switch_to.frame(driver.find_element(By.ID, 'ifrImageWindow'))
                click_here = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, '/html/body/fieldset/object/p/a'))
                )
                driver.execute_script("arguments[0].click();", click_here)
                # 🔽 Add this block right after clicking the download link
                pdf_files_before = set(os.listdir(download_dir))
                time.sleep(10)  # wait for file to appear
                pdf_files_after = set(os.listdir(download_dir))
                new_pdf_files = pdf_files_after - pdf_files_before
                downloaded_filename = next(iter(new_pdf_files), "")

                # Store it in your DataFrame row dictionary
                data[row_index]["Downloaded PDF Filename"] = downloaded_filename
                print(f"✅ PDF Downloaded for row {row_index + 1} on Page {page_number}")
                time.sleep(10)
            except Exception as sub_e:
                print(f"⚠️ Error on row {row_index + 1}: {sub_e}")
            finally:
                try:
                    if len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        driver.close()
                        driver.switch_to.window(original_window)
                except Exception as fe:
                    print(f"⚠️ Cleanup error after row {row_index + 1}: {fe}")
            row_index += 1
    except Exception as e:
        print(f"❌ Failed to process Page {page_number}: {e}")
        break

    # === Pagination: Go to Next Page if available ===
    try:
        next_button = driver.find_element(By.XPATH, '//*[@id="OptionsBar2_imgNext"]')
        if "disabled" in next_button.get_attribute("class") or not next_button.is_displayed():
            print("🚫 No more pages. Pagination ended.")
            break
        else:
            driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
            time.sleep(1)
            next_button.click()
            page_number += 1
            time.sleep(5)  # Wait for next page to load
    except Exception as e:
        print(f"🛑 No 'Next' button found or clickable: {e}")
        break

# After pagination loop finishes
print("✅ All pages processed.")
driver.quit()


def create_dated_directory(base_dir):
    today_str = datetime.now().strftime("Downloaded PDFs %m-%d-%Y")
    dated_path = os.path.join(base_dir, today_str)
    os.makedirs(dated_path, exist_ok=True)
    print(f"Created folder: {dated_path}")
    return dated_path

import re

def wait_for_download(dir_path, timeout=60):
    print("Waiting for file download to complete...")
    seconds = 0
    while seconds < timeout:
        time.sleep(1)
        if not any(fname.endswith(".part") for fname in os.listdir(dir_path)):
            print("Download finished.")
            return
        seconds += 1
    print("Download timed out.")

def remove_duplicate_pdfs(directory):
    print("Starting post-download cleanup...")
    for filename in os.listdir(directory):
        if filename.lower().endswith(".pdf") and "(" in filename:
            file_path = os.path.join(directory, filename)
            try:
                os.remove(file_path)
                print(f"Removed duplicate: {file_path}")
            except Exception as e:
                print(f"Failed to remove {file_path}: {e}")

# 1. Wait for all downloads to finish
wait_for_download(download_dir)

# 2. Remove any duplicate PDFs before moving
remove_duplicate_pdfs(download_dir)

# 3. Create a folder named with today's date
today_str = datetime.now().strftime("Downloaded PDFs %m-%d-%Y")
dest_dir = os.path.join(download_dir, today_str)
os.makedirs(dest_dir, exist_ok=True)
print(f"Created folder: {dest_dir}")

# 4. Move cleaned files into the dated folder
for filename in os.listdir(download_dir):
    if filename.lower().endswith(".pdf"):
        src_path = os.path.join(download_dir, filename)
        dst_path = os.path.join(dest_dir, filename)
        shutil.move(src_path, dst_path)
        print(f"Moved: {filename} → {dst_path}")

# 5. Save the Excel file into the same folder with date in the filename
excel_filename = f"instrument_book_page_results_{datetime.now().strftime('%m-%d-%Y')}.xlsx"
output_path = os.path.join(dest_dir, excel_filename)
df.to_excel(output_path, index=False)
print(f"Excel file saved to: {output_path}")

print("Cleanup and organization complete.")

# Close the driver
driver.quit()

import subprocess
subprocess.run(["python", r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Mcklenburg\pdf_text_parser.py"])

# === Step 2: Locate Updated Excel File ===
base_dir = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Mcklenburg\Scraped File"
folders = [f for f in os.listdir(base_dir) if f.startswith("Downloaded PDFs")]
latest_folder = max(folders, key=lambda f: datetime.strptime(f.replace("Downloaded PDFs ", ""), "%m-%d-%Y"))
pdf_folder = os.path.join(base_dir, latest_folder)
excel_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".xlsx")]
latest_excel = max(excel_files, key=lambda f: os.path.getctime(os.path.join(pdf_folder, f)))
excel_path = os.path.join(pdf_folder, latest_excel)

# === Step 3: Load Excel into DataFrame ===
df = pd.read_excel(excel_path)

# === Step 4: Authenticate and Append to Google Sheets ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open_by_key("1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ").worksheet("Mecklenburg County")

date_scraped = datetime.now().strftime('%m/%d/%Y')

import math

def sanitize_value(val):
    if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
        return ""
    return val

for _, row in df.iterrows():
    new_row = [
        sanitize_value(date_scraped),
        sanitize_value(row.get("Instrument #", "")),
        sanitize_value(row.get("Book - Page", "")),
        sanitize_value(row.get("Date Filed", "")),
        sanitize_value(row.get("Document Type", "")),
        sanitize_value(row.get("Party Name(s)", "")),
        sanitize_value(row.get("Book", "")),
        sanitize_value(row.get("Page", ""))
    ]
    # Find the next available row in Column A only
    def get_next_available_row(sheet, column_letter='A'):
        col_values = sheet.col_values(ord(column_letter.upper()) - 64)  # Column A = 1
        return len(col_values) + 1

    next_row = get_next_available_row(sheet)

    # Update the exact range starting from Column A
    cell_range = f"A{next_row}:H{next_row}"
    sheet.update(cell_range, [new_row], value_input_option="USER_ENTERED")



print("✅ Excel processed and data successfully appended to Google Sheets.")

end_time = time.time()
elapsed_time = end_time - start_time
print(f"\n⏱️ Script completed in {elapsed_time:.2f} seconds.")