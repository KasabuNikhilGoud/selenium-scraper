from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
from concurrent.futures import ThreadPoolExecutor, as_completed
import gspread
import time
from datetime import datetime
from zoneinfo import ZoneInfo

# === CONFIG ===
SHEET_ID = "13ZP7Q9-Yc4mFM64zGosg0CZYWtwWq2RlL9piU2B7qeY"
MAX_ATTEMPTS = 3
MAX_THREADS = 5

# === Google Sheets Setup ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SHEET_ID).sheet1

# === Chrome Options ===
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless=new')
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--disable-notifications')
chrome_options.add_argument('--log-level=3')
chrome_options.add_argument('--window-size=1280,800')

# === Get roll numbers from sheet (row 11+) ===
def get_rolls_from_sheet():
    all_rows = sheet.get_all_values()
    rolls = []
    for row in all_rows[10:]:  # rows after header
        if len(row) > 0 and row[0].strip():
            rolls.append(row[0].strip() + "P")  # add P for login
    return rolls

# === Scraper Worker ===
def process_roll(roll):
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.set_page_load_timeout(10)
            wait = WebDriverWait(driver, 5)

            driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")

            # Username = Password = Roll
            wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(roll)
            driver.find_element(By.ID, "btnNext").click()

            wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(roll)
            driver.find_element(By.ID, "btnSubmit").click()

            # Click Dashboard
            wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()

            # Get Attendance
            wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_lblTotalPercentage")))
            attendance = driver.find_element(By.ID, "ctl00_cpStud_lblTotalPercentage").text.strip()

            print(f"✅ {roll} → {attendance}")
            return (roll, attendance)

        except Exception as e:
            print(f"⚠️ Attempt {attempt} failed: {roll} — {e}")
            time.sleep(0.5)
        finally:
            try:
                driver.quit()
            except:
                pass

    print(f"❌ Max attempts failed: {roll}")
    return (roll, "")

# === Prepare new column (after Name) ===
def prepare_new_column():
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    current_datetime = ist_time.strftime("%Y-%m-%d %I:%M %p")

    # Insert column at position 3 (C)
    sheet.insert_cols([[]], 3)
    sheet.update_cell(10, 3, current_datetime)  # header in row 10

    print(f"📅 Created new column C with timestamp: {current_datetime}")
    return 3  # always C for this run

# === Run Scraping ===
def run_parallel_scraping():
    rolls = get_rolls_from_sheet()
    print(f"📋 Found {len(rolls)} rolls from sheet")

    # Create new column
    col_position = prepare_new_column()

    batch_size = MAX_THREADS
    total_batches = (len(rolls) + batch_size - 1) // batch_size

    for batch_index in range(total_batches):
        start = batch_index * batch_size
        end = start + batch_size
        batch_rolls = rolls[start:end]

        print(f"\n🚀 Batch {batch_index + 1}/{total_batches}: {batch_rolls}")

        batch_results = {}
        with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
            futures = {executor.submit(process_roll, roll): roll for roll in batch_rolls}
            for future in as_completed(futures):
                roll = futures[future]
                try:
                    r, attendance = future.result()
                    batch_results[r] = attendance
                except Exception as e:
                    print(f"❌ Error scraping {roll}: {e}")
                    batch_results[roll] = ""

        # ✅ Update this batch into correct row
        for idx, roll in enumerate(batch_rolls, start=start + 11):  # row 11+
            clean_roll = roll[:-1]  # remove P
            attendance = batch_results.get(roll, "")
            sheet.update_cell(idx, col_position, attendance)

        time.sleep(1)

    print("\n✅ All rolls updated in column C with attendance!")

# === MAIN ===
if __name__ == "__main__":
    run_parallel_scraping()
