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
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# --- Paths ---
firefox_binary_path = r'C:\Program Files\Mozilla Firefox\firefox.exe'
geckodriver_path = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\geckodriver.exe'

# --- Set dynamic dated download folder and clean up old PDFs ---
base_dir = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cleveland\Scraped and Downloads'
today_str = datetime.now().strftime('%m-%d-%Y')
download_dir = os.path.join(base_dir, f"Cleveland Scraped File {today_str}")
os.makedirs(download_dir, exist_ok=True)


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

def append_pdfs_to_google_sheet(download_dir, sheet_id, sheet_name, credentials_path):
    # Set up Google Sheets connection
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(credentials_path, scopes=scope)
    client = gspread.authorize(creds)

    # Open the sheet
    sheet = client.open_by_key(sheet_id).worksheet(sheet_name)

    # Get today's date for column A
    today_str = datetime.now().strftime('%m-%d-%Y')

    # List all PDFs
    pdf_files = [f for f in os.listdir(download_dir) if f.lower().endswith(".pdf")]

    if not pdf_files:
        print("⚠️ No PDFs found to upload to sheet.")
        return

    # Prepare rows
    rows_to_append = [[today_str, filename] for filename in pdf_files]

    # Append all rows
    sheet.append_rows(rows_to_append, value_input_option="USER_ENTERED")
    print(f"✅ Uploaded {len(rows_to_append)} PDF entries to Google Sheet.")

def remove_duplicate_pdfs(download_dir):
    """
    Removes duplicate PDFs from the specified folder.
    Duplicate PDFs typically have names like 'Document (1).pdf', 'Document (2).pdf', etc.
    """
    pattern = re.compile(r"\(\d+\)\.pdf$", re.IGNORECASE)
    removed_files = 0

    for filename in os.listdir(download_dir):
        file_path = os.path.join(download_dir, filename)
        if filename.lower().endswith(".pdf") and pattern.search(filename):
            try:
                os.remove(file_path)
                removed_files += 1
                print(f"🗑️ Removed duplicate: {filename}")
            except Exception as e:
                print(f"⚠️ Could not delete {filename}: {e}")

    print(f"✅ Cleanup complete. Removed {removed_files} duplicate PDF(s).")

# --- Start ---
driver.get("https://us5.courthousecomputersystems.com/ClevelandNCNW/application.asp?resize=true")
driver.maximize_window()
time.sleep(2)

wait = WebDriverWait(driver, 10)
try:
    iframe = wait.until(EC.presence_of_element_located((By.ID, "tabframe0")))
    driver.switch_to.frame(iframe)
except TimeoutException:
    print("Iframe not found on the page.")

# --- Initialize date input for yesterday's and today's date ---

yesterday = '07/01/2025'  # Example static date for testing
#today = '07/15/2025'  # Example static date for testing

#yesterday = (datetime.now() - timedelta(days=1)).strftime('%m/%d/%Y')
today = datetime.now().strftime('%m/%d/%Y')

# Example: Find date input fields and set their values
from_date_input = wait.until(EC.presence_of_element_located((By.ID, 'fromdate')))
to_date_input = wait.until(EC.presence_of_element_located((By.ID, 'todate')))

from_date_input.click()
from_date_input.send_keys(today)
time.sleep(1)  # Wait for the input to register
to_date_input.click()
to_date_input.send_keys(today)
time.sleep(1)  # Wait for the input to register

# --- Wait for the element to be present (not necessarily clickable) ---
wait.until(EC.presence_of_element_located((By.ID, 'exactmatch')))
driver.execute_script("document.getElementById('exactmatch').click();")

# --- Book Type Selection ---
select_element = driver.find_element(By.ID, "availablebooktypes")
book_type_select = Select(select_element)

# Find all REAL PROPERTY matches
matches = [opt for opt in book_type_select.options if opt.text.strip().upper() == 'REAL PROPERTY']

if len(matches) >= 3:
    third_real_property = matches[2]

    # Scroll and click twice
    driver.execute_script("arguments[0].scrollIntoView(true);", third_real_property)
    time.sleep(0.5)
    third_real_property.click()
    time.sleep(0.3)

    # ⏳ Wait for Document Type list to refresh
    time.sleep(2)
else:
    print("❌ Third 'REAL PROPERTY' option not found.")


# --- Document Type Selection ---
doc_type_select = Select(driver.find_element(By.ID, "instrumenttypes"))

# Retry loop just in case the refresh is slow
for _ in range(5):
    options = [opt.text.strip().upper() for opt in doc_type_select.options]
    if 'SUBSTITUTION OF TRUSTEE' in options:
        break
    time.sleep(2)
else:
    print("❌ 'SUBSTITUTION OF TRUSTEE' never appeared.")
    raise Exception("Document Type list did not refresh properly.")

# --- Select 'SUBSTITUTION OF TRUSTEE' Document Type ---
target_doc_type = "SUBSTITUTION OF TRUSTEE"
doc_type_select_element = driver.find_element(By.ID, "instrumenttypes")
doc_type_options = doc_type_select_element.find_elements(By.TAG_NAME, "option")

# Step 1: Deselect all
for option in doc_type_options:
    driver.execute_script("arguments[0].selected = false;", option)

# Step 2: Select only the target type
selected = False
for option in doc_type_options:
    text = option.get_attribute("textContent").strip().upper()
    if text == target_doc_type:
        driver.execute_script("arguments[0].selected = true;", option)
        driver.execute_script("arguments[0].scrollIntoView(true);", option)
        selected = True
        print(f"✅ Selected: {text}")
        break

if not selected:
    print("❌ Could not find 'SUBSTITUTION OF TRUSTEE'")
else:
    # Step 3: Dispatch 'change' event on the full select box
    driver.execute_script("""
        let select = arguments[0];
        let event = new Event('change', { bubbles: true });
        select.dispatchEvent(event);
    """, doc_type_select_element)

    # Step 4: Click the search button
    search_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'search')))
    driver.execute_script("arguments[0].click();", search_button)
    time.sleep(5)

# --- Adjust View Per Page ---
view_per_page_select = Select(driver.find_element(By.ID, "recordsperpage"))
view_per_page_select.select_by_visible_text("500")
time.sleep(2)

# --- Table Scraping and Download ---

results = driver.find_element(By.ID, "resultspane")
if not results:
    print("❌ No results found.")
else:
    print("✅ Results pane found, proceeding to scrape the table.")

    table = results.find_element(By.CLASS_NAME, "results")
    if not table:
        print("❌ No table found in results pane.")
    else:
        print("✅ Table found, proceeding to scrape rows.")

# --- Scroll utility to load all rows ---
def scroll_results_table_to_load_all_rows(driver, container_id="resultspane", delay=0.5, max_scrolls=30):
    container = driver.find_element(By.ID, container_id)
    previous_height = driver.execute_script("return arguments[0].scrollHeight", container)

    for _ in range(max_scrolls):
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", container)
        time.sleep(delay)
        current_height = driver.execute_script("return arguments[0].scrollHeight", container)
        if current_height == previous_height:
            break
        previous_height = current_height

# --- Click "Unique Documents" button ---
try:
    unique_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'span#filterDocumentResultsUniqueDocs img'))
    )
    unique_button.click()
    print("✅ Clicked 'Unique Documents' filter button.")
    time.sleep(3)  # Allow table to refresh
except Exception as e:
    print(f"❌ Failed to click 'Unique Documents' button: {e}")

# --- Scroll to load all unique rows ---
scroll_results_table_to_load_all_rows(driver)
print("✅ Finished scrolling, ready to extract unique rows.")

# --- Grab all rows after unique filter applied ---
table = driver.find_element(By.CLASS_NAME, "results")
rows = table.find_elements(By.TAG_NAME, "tr")[1:]  # skip header row
print(f"✅ Found {len(rows)} unique document rows.")

# --- Function to wait for download to finish ---
def wait_for_download_to_finish(folder, timeout=20):
    end_time = time.time() + timeout
    while time.time() < end_time:
        files = os.listdir(folder)
        if any(f.endswith(".part") for f in files):  # Firefox temp download file
            time.sleep(1)
        elif any(f.lower().endswith(".pdf") for f in files):
            return True
    return False

for i in range(1, len(rows) + 1):
    try:
        # Re-fetch table and rows to avoid stale element errors
        table = driver.find_element(By.CLASS_NAME, "results")
        rows = table.find_elements(By.TAG_NAME, "tr")[1:]  # skip header
        row = rows[i - 1]

        # --- Extract Record Date from column with class "col c4" ---
        try:
            date_cell = row.find_element(By.CSS_SELECTOR, "td.col.c4")
            record_date = date_cell.text.strip()
            print(f"[{i}] 🗓️ Record Date: {record_date}")
        except Exception as e:
            record_date = "N/A"
            print(f"[{i}] ⚠️ Could not extract Record Date: {e}")

        # Click the image icon
        img_icon = row.find_element(By.CSS_SELECTOR, "td.col.c2 img[title='Document image is available']")
        driver.execute_script("arguments[0].scrollIntoView(true);", img_icon)
        time.sleep(0.5)
        img_icon.click()
        print(f"[{i}] ✅ Clicked image for row {i}.")
        time.sleep(4)

        # Switch to iframe with download link
        print(f"[{i}] 🔄 Switching to tabframe1...")
        driver.switch_to.default_content()
        outer_iframe = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "tabframe1"))
        )
        driver.switch_to.frame(outer_iframe)
        print(f"[{i}] ✅ Switched to tabframe1.")

        # Retry logic for inner iframe
        inner_switched = False
        for attempt in range(5):
            print(f"[{i}] ⏳ Attempt {attempt+1}/5: locating inner iframe...")
            try:
                inner_iframe = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'viewimageframe.asp')]"))
                )
                driver.switch_to.frame(inner_iframe)
                print(f"[{i}] ✅ Switched to inner iframe with viewimageframe.asp.")
                inner_switched = True
                break
            except:
                time.sleep(1)

        if not inner_switched:
            raise Exception(f"[{i}] ❌ inner iframe not found after retries.")

        # Locate Download Link
        try:
            download_link = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Download Image')]"))
            )
            print(f"[{i}] 👀 'Download Image' link detected.")
        except Exception as e:
            print(f"[{i}] ❌ Could not locate 'Download Image' link: {e}")
            break

        if not download_link.is_displayed():
            driver.execute_script("arguments[0].scrollIntoView(true);", download_link)
            time.sleep(0.5)

        driver.execute_script("arguments[0].click();", download_link)
        print(f"[{i}] 📥 Clicked 'Download Image' link.")

        # Wait for download
        import random

        # Wait before proceeding to ensure iframe has updated
        delay = random.uniform(6, 10)
        print(f"[{i}] ⏳ Waiting {round(delay, 1)}s for iframe and PDF to fully load...")
        time.sleep(delay)

        # Wait for PDF download to start and finish
        old_files = set(os.listdir(download_dir))

        success = False
        for _ in range(20):  # wait up to 20 seconds
            time.sleep(1)
            new_files = set(os.listdir(download_dir)) - old_files
            pdfs = [f for f in new_files if f.lower().endswith('.pdf')]
            if pdfs:
                success = True
                break

        if success:
            print(f"[{i}] ✅ PDF download completed.")
        else:
            print(f"[{i}] ⚠️ No new PDF detected. Moving on anyway.")

        time.sleep(2)


        time.sleep(5)

        # Get most recent downloaded PDF
        latest_pdf = max(
            [os.path.join(download_dir, f) for f in os.listdir(download_dir) if f.lower().endswith(".pdf")],
            key=os.path.getctime
        )
        pdf_filename = os.path.basename(latest_pdf)


        # Extract Grantors and Grantees
        try:
            # Find all <td> cells that have the 'Grantors:' and 'Grantees:' label
            td_elements = driver.find_elements(By.XPATH, "//td[.//span[text()='Grantors:'] or .//span[text()='Grantees:']]")

            grantors = "N/A"
            grantees = "N/A"

            for td in td_elements:
                label = td.find_element(By.XPATH, ".//span").text.strip().upper()
                divs = td.find_elements(By.XPATH, ".//div")
                names = "; ".join([div.text.strip() for div in divs if div.text.strip()])
                if label == "GRANTORS:":
                    grantors = names
                elif label == "GRANTEES:":
                    grantees = names

            print(f"[{i}] 📝 Grantors: {grantors}")
            print(f"[{i}] 📝 Grantees: {grantees}")

        except Exception as e:
            grantors = "N/A"
            grantees = "N/A"
            print(f"[{i}] ⚠️ Failed to extract Grantors/Grantees: {e}")

        # Append row to Google Sheet
        try:
            sheet = gspread.authorize(Credentials.from_service_account_file(
                r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cleveland\credentials.json",
                scopes=["https://www.googleapis.com/auth/spreadsheets"]
            )).open_by_key("1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ").worksheet("Cleveland County")

            sheet.append_row([datetime.now().strftime('%m-%d-%Y'), record_date, pdf_filename, grantors, grantees], value_input_option="USER_ENTERED")
            print(f"[{i}] ✅ Appended to sheet: {pdf_filename}")
        except Exception as e:
            print(f"[{i}] ❌ Failed to append to sheet: {e}")

        # Return to table view
        try:
            driver.switch_to.default_content()
            print(f"[{i}] 🔁 Returning to Consolidated Index (tab0)...")

            tab_link = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "tab0"))
            )
            driver.execute_script("arguments[0].click();", tab_link)
            print(f"[{i}] 🔁 Clicked tab0")

            table_iframe = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "tabframe0"))
            )
            driver.switch_to.frame(table_iframe)

            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "results"))
            )
            print(f"[{i}] ✅ Table reloaded.")
            time.sleep(1)

        except Exception as e:
            print(f"[{i}] ❌ Failed to return to table view: {e}")
            break

    except Exception as e:
        print(f"[{i}] ❌ Failed on row {i}: {e}")
        break

# --- Final Cleanup ---
remove_duplicate_pdfs(download_dir)

driver.quit()

import subprocess
subprocess.run(["python", r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cleveland\Cleveland_book_and_page.py"])
subprocess.run(["python", r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cleveland\Cleveland_Deed_of_Trust.py"])
subprocess.run(["python", r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cleveland\Cleveland_DOT_Property_Address.py"])
