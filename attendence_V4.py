from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
from concurrent.futures import ThreadPoolExecutor, as_completed
import gspread
import time
from datetime import datetime
from zoneinfo import ZoneInfo  # ✅ Built-in timezone support (Python 3.9+)

# === CONFIG ===
SHEET_ID = "13ZP7Q9-Yc4mFM64zGosg0CZYWtwWq2RlL9piU2B7qeY"
MAX_ATTEMPTS = 3          # retry attempts for each roll
MAX_THREADS = 5           # number of students processed in parallel

# === Google Sheets Setup ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SHEET_ID).sheet1

# === Chrome Options ===
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless=new')  # keep headless
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--disable-notifications')
chrome_options.add_argument('--log-level=3')
chrome_options.add_argument('--window-size=1280,800')

# === Get Roll Numbers from Sheet (row 11+) ===
def get_rolls_from_sheet():
    all_rows = sheet.get_all_values()
    rolls = []
    for row in all_rows[10:]:  # rows after 10 (start from row 11)
        if len(row) > 0 and row[0].strip():  # has roll number
            rolls.append(row[0].strip())    # e.g. 237Z1A0572
    return rolls, all_rows

# === Scraper Worker ===
def process_roll(roll):
    # ERP login requires adding "P"
    erp_roll = roll + "P"

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.set_page_load_timeout(10)
            wait = WebDriverWait(driver, 5)

            driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")

            # Step 1: Enter Username (with P)
            wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(erp_roll)
            driver.find_element(By.ID, "btnNext").click()

            # Step 2: Enter Password (with P)
            wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(erp_roll)
            driver.find_element(By.ID, "btnSubmit").click()

            # Step 3: Click dashboard link
            wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()

            # Step 4: Get attendance %
            wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_lblTotalPercentage")))
            attendance = driver.find_element(By.ID, "ctl00_cpStud_lblTotalPercentage").text.strip()

            print(f"✅ {roll} → {attendance}")
            return (roll, attendance)  # return original roll (without P)

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

# === Create new column ONCE before starting ===
def prepare_new_column(all_rows):
    col_position = 3  # after Name
    
    # ✅ Get IST time using built-in zoneinfo (no pytz needed)
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    current_datetime = ist_time.strftime("%Y-%m-%d %I:%M %p")  # Example: 2025-07-17 08:45 AM
    
    # ✅ Insert new column and add timestamp as header
    sheet.insert_cols([[]], col_position)
    sheet.update_cell(10, col_position, current_datetime)
    print(f"📅 Created new column (IST): {current_datetime}")

    # Create roll → row mapping for quick lookup
    roll_to_row = {}
    for idx, row in enumerate(all_rows[10:], start=11):
        if len(row) >= 1 and row[0].strip():
            roll_to_row[row[0].strip()] = idx

    return col_position, roll_to_row

# === Update Google Sheet ONLY for current batch ===
def update_batch(batch_results, roll_to_row, col_position):
    for roll, attendance in batch_results.items():
        if roll in roll_to_row:
            row_idx = roll_to_row[roll]
            sheet.update_cell(row_idx, col_position, attendance)
        else:
            print(f"⚠️ Roll {roll} not found in sheet!")
    print(f"✅ Updated {len(batch_results)} rows for this batch")

# === Parallel Scraping ===
def run_parallel_scraping():
    rolls, all_rows = get_rolls_from_sheet()
    print(f"📋 Found {len(rolls)} rolls from sheet")

    # Create column + mapping once
    col_position, roll_to_row = prepare_new_column(all_rows)

    batch_size = MAX_THREADS
    total_batches = (len(rolls) + batch_size - 1) // batch_size

    for batch_index in range(total_batches):
        start = batch_index * batch_size
        end = start + batch_size
        batch_rolls = rolls[start:end]

        print(f"\n🚀 Starting batch {batch_index + 1}/{total_batches}: {batch_rolls}")

        # scrape current batch
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

        # ✅ After this batch finishes, update only this batch
        update_batch(batch_results, roll_to_row, col_position)

        time.sleep(1)  # small delay to avoid API quota

    print("\n✅ All batches completed and updated batch by batch!")

# === MAIN ===
if __name__ == "__main__":
    run_parallel_scraping()
