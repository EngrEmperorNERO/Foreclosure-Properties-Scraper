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
from selenium.webdriver.support.ui import Select
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
import os
import gspread
from google.oauth2.service_account import Credentials
import requests


# --- Paths ---
firefox_binary_path = r'C:\Program Files\Mozilla Firefox\firefox.exe'
geckodriver_path = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\geckodriver.exe'
firefox_profile_path = r'C:\Users\Zemo\AppData\Roaming\Mozilla\Firefox\Profiles\gjvhrvid.default-release'

# --- Set dynamic dated download folder and clean up old PDFs ---
base_dir = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Union\E-Courts'
today_str = datetime.now().strftime('%m-%d-%Y')
download_dir = os.path.join(base_dir, f"Union E-Courts Scraped File {today_str}")
os.makedirs(download_dir, exist_ok=True)


# Set Excel output path
output_excel = os.path.join(download_dir, "union_subt_all.xlsx")

options = Options()
profile = FirefoxProfile(firefox_profile_path)
options.profile = profile
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
driver.get("https://portal-nc.tylertech.cloud/Portal/")
driver.maximize_window()
time.sleep(2)

# Scroll down to the bottom of the page
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(2)

#Search Hearings
search_hearings = driver.find_element(By.ID, "portlet-26")
search_hearings.click()
time.sleep(2)

#Select Location
location_select = Select(driver.find_element(By.ID, "cboHSLocationGroup"))
location_select.select_by_visible_text("Union County")
time.sleep(2)

#Select Hearing Type
hearing_type_select = Select(driver.find_element(By.ID, "cboHSHearingTypeGroup"))
hearing_type_select.select_by_visible_text("All Hearing Types")  
time.sleep(2)

#Select Search Type
search_type_select = Select(driver.find_element(By.ID, "cboHSSearchBy"))
search_type_select.select_by_visible_text("Courtroom")   
time.sleep(2)

#Select Courtroom
select_courtroom = Select(driver.find_element(By.ID, "selHSCourtroom"))
select_courtroom.select_by_visible_text("Union Co. Courthouse, Hearing Room 1046")
time.sleep(2)

# Scroll to the bottom of the page to ensure all elements are loaded
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(2)

#yesterday = datetime.now() - timedelta(days=3)
today = datetime.now().strftime('%m/%d/%Y')

yesterday = '07/24/2025'
#today = '07/07/2025'

#Date From
date_from_input = driver.find_element(By.ID, "SearchCriteria_DateFrom")
date_from_input.clear() 
date_from_input.send_keys(today)

#Date To
date_to_input = driver.find_element(By.ID, "SearchCriteria_DateTo")
date_to_input.clear()
date_to_input.send_keys(today)

# Click anywhere to close floating window if present
try:
    body = driver.find_element(By.TAG_NAME, "body")
    body.click()
    time.sleep(1)
except Exception:
    pass

# Click anywhere to close floating window if present
try:
    body = driver.find_element(By.TAG_NAME, "body")
    body.click()
    time.sleep(1)
except Exception:
    pass

print("⏳ Waiting for CAPTCHA to be solved by 2Captcha...")

try:
    WebDriverWait(driver, 180).until(
        EC.text_to_be_present_in_element(
            (By.CLASS_NAME, "ReCaptcha_solver"),
            "SOLVED"
        )
    )
    print("✅ CAPTCHA solved.")

    # Wait for loading spinner to disappear
    WebDriverWait(driver, 30).until(
        EC.invisibility_of_element_located((By.CLASS_NAME, "k-loading-image"))
    )

    # Wait for the DOM to fully load
    WebDriverWait(driver, 30).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

except Exception as e:
    print(f"❌ CAPTCHA solve timeout or page error: {e}")
    driver.quit()
    raise


#Click Submit
submit_button = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, 'btnHSSubmit'))
)
driver.execute_script("arguments[0].click();", submit_button)


# Wait for loading spinner to disappear
WebDriverWait(driver, 30).until(
    EC.invisibility_of_element_located((By.CLASS_NAME, "k-loading-image"))
)

# Click the dropdown safely with retry logic
dropdown_xpath = "//span[contains(@class, 'k-dropdown')]"
max_attempts = 5

for attempt in range(max_attempts):
    try:
        print(f"🔁 Attempt {attempt+1} to click items dropdown...")
        # Wait for dropdown to become clickable
        items_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, dropdown_xpath))
        )

        # Ensure loading spinner is gone
        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "k-loading-image"))
        )

        driver.execute_script("arguments[0].click();", items_dropdown)
        print("✅ Clicked items dropdown.")
        break
    except Exception as e:
        print(f"⚠️ Click failed (attempt {attempt+1}): {e}")
        time.sleep(3)
else:
    print("❌ Failed to click items dropdown after multiple attempts.")
    driver.quit()
    raise Exception("Dropdown click failed after retries.")


# Wait for the popup list to become visible
items_popup_option = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, "//li[@role='option' and text()='200']"))
)

# Click on '200'
driver.execute_script("arguments[0].click();", items_popup_option)
time.sleep(3)

# Wait for loading spinner to disappear before sorting
WebDriverWait(driver, 20).until(
    EC.invisibility_of_element_located((By.CLASS_NAME, "k-loading-image"))
)

# Locate the "Case Type" header
case_type_header = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//th[@data-field='HearingDateParsed']"))
)

# Click twice to sort in ascending order
driver.execute_script("arguments[0].click();", case_type_header)
time.sleep(1)
driver.execute_script("arguments[0].click();", case_type_header)
time.sleep(2)

# Wait again for results to reload after sort
WebDriverWait(driver, 20).until(
    EC.invisibility_of_element_located((By.CLASS_NAME, "k-loading-image"))
)

# Scroll to the top of the page
driver.execute_script("window.scrollTo(0, 0);")
time.sleep(1)

# Get all foreclosure rows first (store references before page reloads)
rows = driver.find_elements(By.XPATH, "//table[@role='table']//tbody/tr")

foreclosure_rows = []
for row in rows:
    try:
        case_type = row.find_element(By.XPATH, ".//td[3]//div[@class='big-search-type']").text.strip()
        if "Foreclosure (Special Proceeding)" in case_type:
            case_link = row.find_element(By.XPATH, ".//a[contains(@class, 'caseLink') and @data-url]")
            foreclosure_rows.append(case_link)
    except:
        continue

print(f"Found {len(foreclosure_rows)} foreclosure case(s) to process.")

party_data = []  # store scraped names for Google Sheets upload

# Now process each foreclosure link one by one
for index, case_link in enumerate(foreclosure_rows, start=1):
    try:
        # Extract case number text before switching context
        case_number_text = case_link.text.strip().replace(".pdf", "")
        print(f"[{index}] Clicking foreclosure case: {case_number_text}")
        
        # Click the link and switch context
        driver.execute_script("arguments[0].click();", case_link)
        time.sleep(5)

        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])
            print("Switched to new tab.")
        
        # ✅ Now use `case_number_text` in your data collection
        time.sleep(10)

        # Switch to new tab if it opens
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])
            print("Switched to new tab.")

        # Wait for the case details page to load
        time.sleep(5)

        #Click the "Party Information" button
        party_info_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/md-sidenav/md-content/div[4]/div/button"))
        )
        party_info_button.click()
        time.sleep(2)

        # Scrape Respondents
        respondent_names = []
        trustee_names = []

        # Find all party sections (each is a tbody block for a connectionType)
        connection_blocks = driver.find_elements(By.XPATH, "//tbody[@ng-repeat='connectionType in roaSection.casePartyConnectionsTypes']")
        print(f"🔍 Found {len(connection_blocks)} party connection blocks.")

        for i, block in enumerate(connection_blocks, start=1):
            try:
                block_html = block.get_attribute("innerHTML").lower()

                # Detect role type
                if "respondent" in block_html:
                    role_type = "respondent"
                    print(f"🧭 Block {i}: Detected Respondent section.")
                elif "substitute trustee" in block_html:
                    role_type = "substitute trustee"
                    print(f"🧭 Block {i}: Detected Substitute Trustee section.")
                else:
                    print(f"⏩ Block {i}: Skipped (unrelated party type).")
                    continue

                # Extract rows under this party type
                rows = block.find_elements(By.XPATH, ".//tr[contains(@class, 'roa-party-row')]")
                print(f"🔹 Found {len(rows)} party rows under {role_type.title()}.")

                for row in rows:
                    try:
                        name_el = row.find_element(By.XPATH, ".//td[2]")  # Second column always holds name block
                        name = name_el.text.strip().split("\n")[0]        # Get first line only (name)
                        if name:
                            if role_type == "respondent":
                                respondent_names.append(name)
                            elif role_type == "substitute trustee":
                                trustee_names.append(name)
                            print(f"✅ Found {role_type.title()}: {name}")
                    except Exception as inner_e:
                        print(f"⚠️ Skipping row due to inner error: {inner_e}")
                        continue

            except Exception as outer_e:
                print(f"⚠️ Skipping block due to outer error: {outer_e}")
                continue

        time.sleep(5)

        # Click the "Case Events" button
        case_events = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div[1]/div/md-sidenav/md-content/div[7]/div/button"))
        )
        case_events.click()
        time.sleep(2)

        # Wait for the Foreclosure NOH label to appear
        label_xpath = "//div[contains(@class, 'roa-label') and contains(text(), 'Foreclosure (Special Proceeding) Notice of Hearing')]"
        foreclosure_label = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, label_xpath))
        )

        # From the label, locate the correct icon in the same row above
        pdf_icon = foreclosure_label.find_element(
            By.XPATH,
            "./ancestor::div[contains(@class, 'event-row') or contains(@class, 'roa-align-top')]//img[contains(@title, 'A document is available')]"
        )

        # Click the matched icon
        driver.execute_script("arguments[0].click();", pdf_icon)
        print("PDF download triggered for Foreclosure NOH.")
        time.sleep(3)

        # ✅ Append data after all scraping & before switching back
        party_data.append([
            datetime.now().strftime('%Y-%m-%d'),  # Date Scraped
            case_number_text,
            "Foreclosure (Special Proceeding)",
            ", ".join(respondent_names),
            ", ".join(trustee_names)
        ])
        print(f"📦 Appended data for {case_number_text}")

        # Close the tab and return to original window
        if len(driver.window_handles) > 1:
            driver.close()  # Close current tab
            driver.switch_to.window(driver.window_handles[0])
            print("Closed tab and switched back to main window.")
            time.sleep(5)

    except Exception as e:
        print(f"[{index}] Error processing case: {e}")
        try:
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
        except:
            pass
        continue


# --- Upload to Google Sheet (append instead of overwrite) ---
try:
    print("Uploading case details with party info to Google Sheet...")

    # Credentials & Sheet Info
    credentials_path = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Union\E-Courts\credentials.json"
    sheet_id = "1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ"
    sheet_name = "Union County"

    # Authenticate
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file(credentials_path, scopes=scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id).worksheet(sheet_name)

    # Find the next available row
    existing_records = sheet.get_all_values()
    next_row = len(existing_records) + 1  # row after last filled

    # Headers only written once (if empty)
    if next_row == 1:
        headers = ["Date Scraped", "Case Number", "Case Type", "Respondents", "Substitute Trustees"]
        sheet.update(range_name="A1:E1", values=[headers])
        next_row = 2

    # Upload data starting at next available row
    sheet.update(range_name=f"A{next_row}:E{next_row + len(party_data) - 1}", values=party_data)

    print(f"✅ Uploaded {len(party_data)} row(s) to Google Sheet starting at row {next_row}.")

except Exception as e:
    print(f"❌ Failed to upload to Google Sheet: {e}")

driver.quit()

import subprocess

subprocess.run(["python", r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Union\E-Courts\Union_Parser.py"])
