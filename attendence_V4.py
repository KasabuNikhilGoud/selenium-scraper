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

BASE_PREFIX = "237Z1A05"  # Fixed part before the roll number sequence

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

# === Generate Roll Numbers ===
def generate_roll_numbers():
    rolls = []

    # Phase 1: numeric sequence (72–99)
    for num in range(72, 100):
        rolls.append(f"{BASE_PREFIX}{num}P")

    # Phase 2: After 99 → A0, A1…A9, B0…B9, ...
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    for letter in letters:
        for d in range(0, 10):
            code = f"{letter}{d}"
            # Skip invalids
            if code in ["A0", "88", "80"]:
                continue
            rolls.append(f"{BASE_PREFIX}{code}P")

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

# === Prepare Column Header & Auto-expand columns ===
def prepare_new_column():
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    current_datetime = ist_time.strftime("%Y-%m-%d %I:%M %p")

    # ✅ Get sheet size
    rows = sheet.row_count
    cols = sheet.col_count

    # ✅ Find next available column after header row
    headers = sheet.row_values(10)  # Row 10 is header
    col_position = len(headers) + 1

    # ✅ Expand sheet if not enough columns
    if col_position > cols:
        needed = col_position - cols
        sheet.add_cols(needed)
        print(f"➕ Added {needed} extra column(s), total now: {cols + needed}")

    # ✅ Finally write timestamp header
    sheet.update_cell(10, col_position, current_datetime)
    print(f"📅 Created new column at col {col_position}: {current_datetime}")
    return col_position

# === Get existing rolls from sheet for quick lookup ===
def get_existing_rolls():
    all_rows = sheet.get_all_values()
    roll_map = {}  # roll → row number
    for idx, row in enumerate(all_rows[10:], start=11):  # after header row
        if len(row) > 0 and row[0].strip():
            roll_map[row[0].strip()] = idx
    return roll_map

# === Auto-expand rows if needed ===
def ensure_row_capacity(required_rows):
    current_rows = sheet.row_count
    if required_rows > current_rows:
        needed = required_rows - current_rows
        sheet.add_rows(needed)
        print(f"➕ Added {needed} extra rows, total now: {current_rows + needed}")

# === Update Google Sheet ===
def update_sheet(roll, attendance, col_position, roll_map):
    clean_roll = roll[:-1]  # remove trailing 'P' for storage

    if clean_roll in roll_map:
        # Existing student → update in same row
        row_idx = roll_map[clean_roll]
        sheet.update_cell(row_idx, col_position, attendance)
    else:
        # New student → append safely
        print(f"➕ Adding new roll to sheet: {clean_roll}")
        last_row_index = len(sheet.get_all_values()) + 1

        # Ensure enough rows before writing
        ensure_row_capacity(last_row_index)

        # Append new row
        sheet.append_row([clean_roll, "", attendance])  # Roll | Name(empty) | Attendance
        # Update roll_map dynamically
        roll_map[clean_roll] = last_row_index

# === Run Scraping ===
def run_parallel_scraping():
    rolls = generate_roll_numbers()
    print(f"📋 Generated {len(rolls)} roll numbers")

    # Prepare new column header
    col_position = prepare_new_column()
    roll_map = get_existing_rolls()

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

        # ✅ Update each roll result immediately
        for roll, attendance in batch_results.items():
            update_sheet(roll, attendance, col_position, roll_map)

        time.sleep(1)

    print("\n✅ All batches completed & sheet updated!")

# === MAIN ===
if __name__ == "__main__":
    run_parallel_scraping()
