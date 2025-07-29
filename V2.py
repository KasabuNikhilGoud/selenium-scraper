import time
import gspread
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
from shutil import which
from google.oauth2.service_account import Credentials

# Constants
SPREADSHEET_ID = '1Rk3eNqEhbuDIgu3Zx4_CwOZCnFlLm6Vr9obVzYl_zr4'
SHEET_NAME = 'Attendence CSE-B(2023-27)'
TARGET_RANGE = 'D8:D21'
MAX_ATTEMPTS = 2

# Google Sheets setup
creds = Credentials.from_service_account_file("credentials.json", scopes=[
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
])
client = gspread.authorize(creds)
sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
roll_numbers = sheet.get(TARGET_RANGE)
rolls = [r[0].strip() for r in roll_numbers if r]

# Column insert for timestamp header
timestamp = datetime.now().strftime("%d-%b %I:%M %p")
header_row = sheet.row_values(7)
sheet.update_cell(7, len(header_row) + 1, timestamp)


def process_roll(roll):
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            # Chrome setup
            chrome_options = Options()
            chrome_options.add_argument("--headless=new")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--window-size=1920,1080")
            chrome_options.binary_location = which("chromium-browser")

            driver = webdriver.Chrome(options=chrome_options)
            driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")
            wait = WebDriverWait(driver, 10)

            # Step 1: Enter username
            wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(roll)
            driver.find_element(By.ID, "btnNext").click()

            # Step 2: Enter password
            wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(roll)
            driver.find_element(By.ID, "btnSubmit").click()

            # Step 3: Dashboard link
            wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Click Here to go Student Dashbord"))).click()

            # Step 4: Wait for Subject Table
            wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_grdSubject")))
            table = driver.find_element(By.ID, "ctl00_cpStud_grdSubject")
            rows = table.find_elements(By.TAG_NAME, "tr")

            total_classes = 0
            subjects_counted = 0
            for row in rows[1:]:  # skip header
                cols = row.find_elements(By.TAG_NAME, "td")
                if len(cols) >= 4:
                    held = cols[3].text.strip()
                    if held.isdigit():
                        total_classes += int(held)
                        subjects_counted += 1

            avg_classes = round(total_classes / subjects_counted, 1) if subjects_counted else "NA"

            # Insert value to correct row
            cell = sheet.find(roll)
            sheet.update_cell(cell.row, len(header_row) + 1, avg_classes)
            driver.quit()
            return

        except Exception as e:
            print(f"[{roll}] Attempt {attempt} failed: {e}")
            if attempt == MAX_ATTEMPTS:
                try:
                    cell = sheet.find(roll)
                    sheet.update_cell(cell.row, len(header_row) + 1, "Error")
                except:
                    pass
            if 'driver' in locals():
                driver.quit()
            time.sleep(1)


# Run in parallel
with ThreadPoolExecutor(max_workers=3) as executor:
    executor.map(process_roll, rolls)

print("✅ Done inserting class held counts per roll number.")
