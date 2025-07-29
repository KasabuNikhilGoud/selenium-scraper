from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import time
import subprocess
from shutil import which
import tempfile
import shutil
import os

# Setup Chrome options
chrome_options = Options()
# chrome_options.add_argument("--headless=new")  # Uncomment for headless run
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument("--window-size=1280,800")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--no-default-browser-check")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--disable-dev-tools")

# ✅ Fix: Use unique temp user-data-dir to avoid conflicts
user_data_dir = tempfile.mkdtemp()
chrome_options.add_argument(f"--user-data-dir={user_data_dir}")

# Detect Chromium binary
chrome_path = which("chromium-browser")
if chrome_path:
    print(f"✅ Found Chromium at: {chrome_path}")
    chrome_options.binary_location = chrome_path
    try:
        version_output = subprocess.check_output([chrome_path, "--version"], text=True)
        print(f"ℹ️ Chromium version: {version_output.strip()}")
    except subprocess.CalledProcessError as e:
        print(f"⚠️ Failed to get Chromium version: {e}")
else:
    print("⚠️ Chromium not found, using default Chrome")

# Initialize WebDriver
try:
    driver = webdriver.Chrome(options=chrome_options)
    print("✅ WebDriver session initialized")
except Exception as e:
    print(f"⚠️ Failed to initialize WebDriver session: {e}")
    shutil.rmtree(user_data_dir, ignore_errors=True)
    raise

wait = WebDriverWait(driver, 60)

try:
    # Step 1: Open BeeSERP Login Page
    driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")
    print("✅ Opened login page")

    # Step 2: Enter Hall Ticket Number and Click Next
    username_field = wait.until(EC.presence_of_element_located((By.ID, "txtUserName")))
    username_field.send_keys("237Z1A0575P")
    driver.find_element(By.ID, "btnNext").click()
    print("✅ Entered username and clicked Next")

    # Step 3: Enter Password and Click Submit
    password_field = wait.until(EC.presence_of_element_located((By.ID, "txtPassword")))
    password_field.send_keys("237Z1A0575P")
    driver.find_element(By.ID, "btnSubmit").click()
    print("✅ Entered password and clicked Submit")

    # Step 4: Click on Dashboard Link
    dashboard_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Click Here to go Student Dashbord")))
    dashboard_link.click()
    print("✅ Clicked on Student Dashboard")

    # Step 5: Extract Subject-wise Attendance
    classes_held_list = []
    total_held = 0
    max_retries = 5

    for attempt in range(max_retries):
        try:
            wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_PanelSubjectwise")))
            wait.until(EC.visibility_of_element_located((By.XPATH, "//table[@id='ctl00_cpStud_grdSubject']//tr")))
            time.sleep(10)  # Give time for dynamic load

            table_div = driver.find_element(By.ID, "ctl00_cpStud_PanelSubjectwise")
            rows = table_div.find_elements(By.XPATH, ".//tr")[1:]

            for i, row in enumerate(rows, 1):
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) >= 6:
                    subject_name = cells[1].text.strip()
                    held_value = cells[3].text.strip()
                    print(f"Row {i}: Subject: {subject_name}, Classes Held: {held_value}")
                    try:
                        value = int(held_value)
                        if subject_name != "-":
                            classes_held_list.append(value)
                        else:
                            total_held = value
                    except ValueError:
                        classes_held_list.append(0)

            print("✅ Extracted Classes Held:", classes_held_list)
            print(f"✅ Total Classes Held: {total_held}")
            if len(classes_held_list) != 13:
                print(f"⚠️ Warning: Expected 13 rows, got {len(classes_held_list)}")
            break
        except Exception as e:
            print(f"⚠️ Attempt {attempt + 1}/{max_retries} failed: {e}")
            time.sleep(10)
            if attempt == max_retries - 1:
                print("⚠️ Final fallback: filling with zeros")
                classes_held_list = [0] * 13

    # Save HTML snapshot for debugging
    with open("page_source.html", "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    print("✅ Saved page source to page_source.html")

finally:
    driver.quit()
    shutil.rmtree(user_data_dir, ignore_errors=True)

# Step 6: Update Google Sheet (D8:D21)
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

try:
    sheet = client.open_by_key("1Rk3eNqEhbuDIgu3Zx4_CwOZCnFlLm6Vr9obVzYl_zr4").worksheet("Attendence CSE-B(2023-27)")

    # Ensure exactly 14 values
    while len(classes_held_list) < 14:
        classes_held_list.append(0)
    classes_held_list = classes_held_list[:14]

    update_range = 'D8:D21'
    data = [[val] for val in classes_held_list]
    sheet.update(values=data, range_name=update_range)
    print("✅ Classes Held inserted into D8:D21")
except gspread.exceptions.APIError as api_err:
    print(f"⚠️ Google Sheets API Error: {api_err}")
except Exception as e:
    print(f"⚠️ Failed to update Google Sheet: {e}")
