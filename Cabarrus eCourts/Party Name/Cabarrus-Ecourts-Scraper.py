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
from selenium.common.exceptions import UnexpectedAlertPresentException
from datetime import datetime


# --- Paths ---
firefox_binary_path = r'C:\Program Files\Mozilla Firefox\firefox.exe'
geckodriver_path = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\geckodriver.exe'
firefox_profile_path = r'C:\Users\Zemo\AppData\Roaming\Mozilla\Firefox\Profiles\gjvhrvid.default-release'

# --- Set dynamic dated download folder and clean up old PDFs ---
base_dir = r'C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cabarrus eCourts\Party Name'
today_str = datetime.now().strftime('%m-%d-%Y')
download_dir = os.path.join(base_dir, f"Cabarrus E-Courts Scraped File {today_str}")
os.makedirs(download_dir, exist_ok=True)


# Set Excel output path
output_excel = os.path.join(download_dir, "cabarrus_subt_all.xlsx")

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
scraped_date = datetime.now().strftime("%m/%d/%Y")

# --- Start ---
driver.get("https://portal-nc.tylertech.cloud/Portal/")
driver.maximize_window()
time.sleep(2)

# Scroll down to the bottom of the page
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(2)

# Click on the "Smart Search" button
smart_search = driver.find_element(By.ID, "portlet-29")
smart_search.click()
time.sleep(2)

#Locate Advanced Search button and click it
advanced_search = driver.find_element(By.ID, "AdvOptions")  
advanced_search.click()
time.sleep(2)

#General Options
general_options = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div/div/div/form[1]/div/div[2]/div/nav/div/ul/li[2]/a")
general_options.click()
time.sleep(2)

#Scroll down to General Options
driver.execute_script("window.scrollBy(0, 200);")
time.sleep(1)

#Filter by Location
location_filter = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[2]/div/div/div/div/form[1]/div/div[1]/div[2]/div[1]/div/div/div[1]/div/div/button"))
)
location_filter.click()
time.sleep(2)

# Check "Cabarrus County" via JavaScript click on label
try:
    cabarrus_label = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//label[@title="Cabarrus County"]'))
    )
    driver.execute_script("arguments[0].click();", cabarrus_label)
    time.sleep(1)
    print("✅ 'Cabarrus County' checked.")
except Exception as e:
    print(f"❌ Failed to check 'Cabarrus County': {e}")

# THEN uncheck "All Locations (default)" via JavaScript click on label
try:
    all_locations_label = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//label[@title="All Locations (default)"]'))
    )
    driver.execute_script("arguments[0].click();", all_locations_label)
    time.sleep(1)
    print("✅ 'All Locations' unchecked.")
except Exception as e:
    print(f"❌ Failed to uncheck 'All Locations': {e}")


#Scroll down to General Options
driver.execute_script("window.scrollBy(0, 500);")
time.sleep(1)

#Filter by Case Type
case_type_filter = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div/div/div/form[1]/div/div[1]/div[2]/div[3]/div/div/div/fieldset[1]/span/span/input")
case_type_filter.send_keys("Special Proceedings (non-confidential)")
time.sleep(2)

#Filter by Case Status
case_status_filter = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div/div/div/form[1]/div/div[1]/div[2]/div[3]/div/div/div/fieldset[2]/span/span/input")
case_status_filter.send_keys("Pending")
time.sleep(2)

#yesterday = (datetime.now() - timedelta(days=3)).strftime('%m/%d/%Y')
today = datetime.now().strftime('%m/%d/%Y')

yesterday = '05/01/2025'
#today = '07/07/2025'

#Date From
date_from_input = driver.find_element(By.ID, "caseCriteria.FileDateStart")
date_from_input.clear() 
date_from_input.send_keys(today)

#Date To
date_to_input = driver.find_element(By.ID, "caseCriteria.FileDateEnd")
date_to_input.clear()
date_to_input.send_keys(today)

# Scroll back to the top of the page
driver.execute_script("window.scrollTo(0, 0);")
time.sleep(1)

#Back to Search Criteria, use google sheets data as input
# --- Google Sheets auth ---
scope = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
creds = Credentials.from_service_account_file(
    r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cabarrus eCourts\credentials.json",
    scopes=scope
)
client = gspread.authorize(creds)
sheet = client.open_by_key("1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ").worksheet("Lawyers and Lawfirms")
lawyers = sheet.col_values(1)[1:]

results = []

# --- Main loop ---
for i, lawyer_name in enumerate(lawyers, start=1):
    print(f"🔍 [{i}] Searching for: {lawyer_name}")

    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "caseCriteria_SearchCriteria"))).clear()
        driver.find_element(By.ID, "caseCriteria_SearchCriteria").send_keys(lawyer_name)
        time.sleep(5)

        try:
            print("⏳ Waiting for CAPTCHA to be solved...")
            WebDriverWait(driver, 180).until(
                lambda d: "SOLVED" in d.find_element(By.CLASS_NAME, "ReCaptcha_solver").text
            )
            print("✅ CAPTCHA solved!")
        except UnexpectedAlertPresentException:
            alert = driver.switch_to.alert
            alert.dismiss()
            driver.execute_script("document.body.click();")
            time.sleep(2)
        except Exception as e:
            print(f"❌ CAPTCHA not solved: {e}")

        driver.execute_script("arguments[0].click();", driver.find_element(By.ID, "btnSSSubmit"))
        time.sleep(60)

        try:
            case_links = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "caseLink")))
            print(f"🔗 Found {len(case_links)} case link(s).")
            main_window = driver.current_window_handle

            for index, link in enumerate(case_links):
                try:
                    print(f"➡️ Opening case link #{index + 1}")
                    driver.execute_script("arguments[0].click();", link)
                    WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
                    new_window = [w for w in driver.window_handles if w != main_window][0]
                    driver.switch_to.window(new_window)
                    time.sleep(20)

                    try:
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[1]/div/md-sidenav/md-content/div[4]/div/button"))).click()
                        time.sleep(2)
                    except Exception:
                        print("⚠️ Party Information button not found.")

                    # --- Scrape parties ---
                    respondent_names = []
                    trustee_names = []
                    blocks = driver.find_elements(By.XPATH, "//tbody[@ng-repeat='connectionType in roaSection.casePartyConnectionsTypes']")
                    for j, block in enumerate(blocks, start=1):
                        try:
                            block_html = block.get_attribute("innerHTML").lower()
                            if "respondent" in block_html:
                                role_type = "respondent"
                            elif "substitute trustee" in block_html:
                                role_type = "substitute trustee"
                            else:
                                continue

                            rows = block.find_elements(By.XPATH, ".//tr[contains(@class, 'roa-party-row')]")
                            for row in rows:
                                try:
                                    name = row.find_element(By.XPATH, ".//td[2]").text.strip().split("\n")[0]
                                    if role_type == "respondent":
                                        respondent_names.append(name)
                                    elif role_type == "substitute trustee":
                                        trustee_names.append(name)
                                except:
                                    continue

                        except:
                            continue

                    try:
                        case_number_el = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//span[@ng-bind='::data.roaSections.caseSummary.headerInfo.CaseNumber']"))
                        )
                        case_number = case_number_el.text.strip()
                        print(f"📎 Case Number Extracted: {case_number}")
                        # Save filename for linking
                        # Try to find the most recently downloaded file that contains the case number
                        matching_files = [f for f in os.listdir(download_dir) if f.endswith('.pdf') and case_number in f]
                        if matching_files:
                            pdf_filename = sorted(matching_files, key=lambda x: os.path.getmtime(os.path.join(download_dir, x)), reverse=True)[0]
                        else:
                            pdf_filename = f"{case_number}.pdf"  # Fallback
                    except Exception as e:
                        print(f"⚠️ Could not extract case number from page: {e}")
                        case_number = "Unknown"

                    # ✅ Move this out of the above 'except'
                    try:
                        record_date_el = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'ng-binding') and contains(text(), '/') and contains(text(), '202')]"))
                        )
                        record_date = record_date_el.text.strip()
                        print(f"🗓 Record Date Extracted: {record_date}")
                    except Exception as e:
                        print(f"⚠️ Could not extract Record Date: {e}")
                        record_date = ""

                    # --- STEP: Click on Case Events tab ---
                    # ✅ Case Events Tab
                    try:
                        case_events_button = WebDriverWait(driver, 20).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[.//span[contains(text(),'Case Events')]]"))
                        )
                        driver.execute_script("arguments[0].click();", case_events_button)
                        print("📁 Clicked 'Case Events' tab")
                        time.sleep(3)
                    except Exception as e:
                        print(f"⚠️ Could not click 'Case Events': {e}")

                    time.sleep(10)

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


                    for name in respondent_names:
                        results.append({
                            "Lawyer Name": lawyer_name,
                            "Case Number": case_number,
                            "Record Date": record_date,
                            "Scraped Date": scraped_date,
                            "Party Type": "Respondent",
                            "Party Name": name,
                            "Notice PDF": pdf_filename
                        })

                    for name in trustee_names:
                        results.append({
                            "Lawyer Name": lawyer_name,
                            "Case Number": case_number,
                            "Record Date": record_date,
                            "Scraped Date": scraped_date,
                            "Party Type": "Substitute Trustee",
                            "Party Name": name,
                            "Notice PDF": pdf_filename
                        })


                    # --- Remove duplicate PDFs that contain (1), (2), etc. ---
                    for fname in os.listdir(download_dir):
                        if fname.lower().endswith(".pdf") and "(" in fname:
                            try:
                                os.remove(os.path.join(download_dir, fname))
                                print(f"🗑 Removed duplicate file: {fname}")
                            except Exception as e:
                                print(f"⚠️ Could not delete file {fname}: {e}")

                    driver.close()
                    driver.switch_to.window(main_window)
                    time.sleep(3)

                except Exception as e:
                    print(f"⚠️ Failed to process case #{index + 1}: {e}")
                    driver.switch_to.window(main_window)
                    continue

        except Exception as e:
            print(f"❌ Case link error: {e}")

    except Exception as e:
        print(f"⚠️ Error for lawyer {lawyer_name}: {e}")

# --- Final export ---
import gspread
from google.oauth2.service_account import Credentials

# --- Google Sheets Auth ---
scope = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file(
    r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Cabarrus eCourts\credentials.json",
    scopes=scope
)
client = gspread.authorize(creds)
sheet = client.open_by_key("1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ")
worksheet = sheet.worksheet("Cabarrus eCourts Lawyers")

# --- Format data ---
if results:
    df = pd.DataFrame(results)

    formatted_df = df.pivot_table(
    index=["Lawyer Name", "Case Number", "Record Date", "Scraped Date", "Notice PDF"],
    columns="Party Type",
    values="Party Name",
    aggfunc=lambda x: " & ".join(pd.unique(x))
    ).reset_index()

    # --- Load existing data to avoid duplicates ---
    existing_records = worksheet.get_all_values()[1:]  # Skip header
    existing_set = set((row[0], row[1]) for row in existing_records)  # Lawyer Name + Case Number

    # Filter new rows
    new_rows = []
    for _, row in formatted_df.iterrows():
        key = (row["Lawyer Name"], row["Case Number"])
        if key not in existing_set:
            new_rows.append([
            row.get("Lawyer Name", ""),
            row.get("Case Number", ""),
            row.get("Record Date", ""),
            row.get("Scraped Date", ""),
            row.get("Notice PDF", ""),
            row.get("Respondent", ""),
            row.get("Substitute Trustee", "")
        ])

    # Append new rows
    if new_rows:
        worksheet.append_rows(new_rows)
        print(f"✅ Appended {len(new_rows)} new rows to Google Sheet.")
    else:
        print("ℹ️ No new rows to append — all entries already exist.")
else:
    print("⚠️ No results to export.")

driver.quit()
print("✅ Scraping completed and data exported to Google Sheets.")