import os
import time
import zipfile
import glob
import shutil
from datetime import datetime
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials


from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

start_time = time.time()
# ---------------- SETUP ----------------
firefox_binary_path = r'C:\Program Files\Mozilla Firefox\firefox.exe'
geckodriver_path = r'C:\Users\Zemo\Desktop\State&Liberty\Driver\geckodriver.exe'
base_dir = r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Mcklenburg\Scraped File"

# Google Sheets setup
sheet_id = "1C6Q6iJTzO89LJRw6q2K1V-9m8NCzWegHgswfjPHanAQ"
sheet_name = "Mecklenburg County"
credentials_path = "credentials.json"

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file(credentials_path, scopes=scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(sheet_id).worksheet(sheet_name)


# --- Find the latest folder ---
folders = [os.path.join(base_dir, f) for f in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, f))]
latest_folder = max(folders, key=os.path.getmtime)

# --- Find latest Excel file in that folder ---
excel_files = glob.glob(os.path.join(latest_folder, "instrument_book_page_results_*.xlsx"))
latest_excel = max(excel_files, key=os.path.getmtime)

# --- Set up download directory inside the folder ---
today_str = datetime.today().strftime('%m-%d-%Y')
final_pdf_dir = os.path.join(latest_folder, f"Final PDF ({today_str})")
os.makedirs(final_pdf_dir, exist_ok=True)

# --- Load Excel ---
df = pd.read_excel(latest_excel)

# --- Firefox setup ---
options = Options()
options.binary_location = firefox_binary_path
options.set_preference("browser.download.folderList", 2)
options.set_preference("browser.download.dir", final_pdf_dir)
options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/pdf")
options.set_preference("pdfjs.disabled", True)
options.set_preference("browser.download.manager.showWhenStarting", False)
options.set_preference("browser.download.panel.shown", False)

# WebDriver
service = Service(geckodriver_path)
driver = webdriver.Firefox(service=service, options=options)
wait = WebDriverWait(driver, 20)

# --- Navigate to website ---
driver.get("https://meckrod.manatron.com/")
driver.maximize_window()

# Accept terms
wait.until(EC.element_to_be_clickable((By.ID, "cph1_lnkAccept"))).click()

# Real Estate Button
wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div[3]/div[1]/div[2]/div/ul/li[13]/a'))).click()

# Search Real Estate Index
wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div[3]/div[1]/div[2]/div[2]/div/div/div/ul/li[3]/a'))).click()


# --- Main loop for Book/Page search and download ---
results_table_xpath = '/html/body/div[3]/form/div[3]/div[3]/table/tbody/tr[4]/td/div/div/table/tbody/tr[2]/td/table/tbody[2]'

for index, row in df.iterrows():
    try:
        book = str(row['Book'])
        page = str(row['Page'])

        # Wait for fields to be ready
        wait.until(EC.presence_of_element_located((By.ID, "cphNoMargin_f_txtBook"))).clear()
        driver.find_element(By.ID, "cphNoMargin_f_txtBook").send_keys(book)
        driver.find_element(By.ID, "cphNoMargin_f_txtPage").clear()
        driver.find_element(By.ID, "cphNoMargin_f_txtPage").send_keys(page)

        # Click Search
        # Click Search
        driver.find_element(By.ID, "cphNoMargin_SearchButtons2_btnSearch__3").click()
        time.sleep(3)

        # Check if zero results were returned
        try:
            total_rows = driver.find_element(By.ID, "cphNoMargin_cphNoMargin_SearchCriteriaTop_TotalRows").text.strip()
            if total_rows == "0":
                print(f"❌ No records found for Book: {book}, Page: {page}. Skipping.")
                
                # Return to search page
                wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div[3]/div[1]/div[2]/div[1]/ul/li[13]/a'))).click()
                wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div[3]/div[1]/div[2]/div[2]/div/div/div/ul/li[3]/a'))).click()
                continue
        except Exception as e_zero:
            print(f"⚠️ Could not check for record count: {e_zero}")

        try:
            # Wait to see if results table or "no results" message appears
            time.sleep(2)
            if "No results found" in driver.page_source or "No documents matched" in driver.page_source:
                print(f"❌ No results found for Book: {book}, Page: {page}. Returning to search.")
                # Return to search page
                wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div[3]/div[1]/div[2]/div[1]/ul/li[13]/a'))).click()
                wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div[3]/div[1]/div[2]/div[2]/div/div/div/ul/li[3]/a'))).click()
                continue  # Skip to next row

            # Locate results table
            results_table = wait.until(EC.presence_of_element_located((By.XPATH, results_table_xpath)))
            first_view = results_table.find_element(By.CLASS_NAME, "staticLink")

            
            # Click the first 'View' link
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", first_view)
            time.sleep(0.5)
            original_window = driver.current_window_handle
            first_view.click()
            print(f"Clicked 'View' for Book: {book}, Page: {page}")

            WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
            new_window = [w for w in driver.window_handles if w != original_window][0]
            driver.switch_to.window(new_window)
            print("Switched to pop-up")

            try:
                WebDriverWait(driver, 10).until(
                    EC.frame_to_be_available_and_switch_to_it((By.ID, 'ImageViewer1_ifrLTViewer'))
                )
                print("Switched to image iframe")
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.ID, 'form1'))
                )
                get_image_button = WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.ID, 'btnProcessNow'))
                )
                driver.execute_script("arguments[0].click();", get_image_button)
                print("Clicked 'Get Image Now'")
                time.sleep(30)
                WebDriverWait(driver, 20).until(
                    EC.visibility_of_element_located((By.ID, "dlgImageWindow"))
                )
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.ID, 'ifrImageWindow'))
                )
                driver.switch_to.frame(driver.find_element(By.ID, 'ifrImageWindow'))
                print("Switched to image iframe")
                click_here = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, '/html/body/fieldset/object/p/a'))
                )
                driver.execute_script("arguments[0].click();", click_here)
                print("Clicked 'here' link")
                print("Waiting for file download to complete...")
                time.sleep(10)
                print("Download finished.")

                # Append latest downloaded PDF to Google Sheet
                # --- Find the latest downloaded PDF
                downloaded_files = [f for f in os.listdir(final_pdf_dir) if f.lower().endswith(".pdf")]
                latest_pdf = max(downloaded_files, key=lambda f: os.path.getmtime(os.path.join(final_pdf_dir, f)))

                # --- Find the matching row in Google Sheets by Book & Page
                g_sheet_data = sheet.get_all_values()
                header = g_sheet_data[0]
                book_col = header.index("Book")
                page_col = header.index("Page")
                pdf_col = header.index("PDF Filename")

                for i, row_data in enumerate(g_sheet_data[1:], start=2):  # start=2 for 1-based index + header
                    try:
                        sheet_book = str(row_data[book_col]).strip()
                        sheet_page = str(row_data[page_col]).strip()

                        if sheet_book == book and sheet_page == page:
                            sheet.update_cell(i, pdf_col + 1, latest_pdf)
                            print(f"✅ Updated row {i} in Google Sheet with: {latest_pdf}")
                            break
                    except Exception as gs_error:
                        print(f"❌ Error updating sheet at row {i}: {gs_error}")

            except Exception as sub_e:
                print(f"Error inside iframe for Book: {book}, Page: {page}: {sub_e}")
            finally:
                try:
                    driver.switch_to.default_content()
                    driver.close()
                    driver.switch_to.window(original_window)
                    print("Closed pop-up window.")
                    print("Returned to main window\n")
                except Exception as close_e:
                    print(f"Cleanup error: {close_e}")
        except Exception as view_e:
            print(f"No viewable result found for Book: {book}, Page: {page}. Error: {view_e}")
            continue

        # Return to search page
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div[3]/div[1]/div[2]/div[1]/ul/li[13]/a'))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/form/div[3]/div[1]/div[2]/div[2]/div/div/div/ul/li[3]/a'))).click()


    except Exception as e:
        print(f"Failed processing Book: {book}, Page: {page}. Error: {e}")
        continue

print("All downloads initiated. Closing browser.")
driver.quit()

import subprocess
subprocess.run(["python", r"C:\Users\Zemo\Desktop\Atlas Residential\Scraper\Mcklenburg\pdf_text_parser_address_with_OCR_Logs.py"])

end_time = time.time()
elapsed_time = end_time - start_time
print(f"\n⏱️ Script completed in {elapsed_time:.2f} seconds.")