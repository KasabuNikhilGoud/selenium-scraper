from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
from concurrent.futures import ThreadPoolExecutor, as_completed
from shutil import which
import gspread
import time
from datetime import datetime
from zoneinfo import ZoneInfo
from bs4 import BeautifulSoup

# === CONFIG ===
SHEET_ID = "13ZP7Q9-Yc4mFM64zGosg0CZYWtwWq2RlL9piU2B7qeY"
MAX_ATTEMPTS = 3
MAX_THREADS = 15
BASE_PREFIX = "237Z1A05"
SUBJECTS = [
    "DAA", "CN", "DEVOPS", "PPL", "NLP", "CN LAB", "DEVOPS LAB",
    "ACS LAB", "IPR", "Sports", "Mentoring", "Association", "Library"
]

# === Google Sheets Setup ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(SHEET_ID)
sheet1 = spreadsheet.worksheet("Attendence CSE-B(2023-27)")

# === Setup Chrome Options ===
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument("--window-size=1280,800")

chrome_path = which("chromium-browser")
if chrome_path:
    print(f"✅ Found Chromium at: {chrome_path}")
    chrome_options.binary_location = chrome_path
else:
    print("⚠️ Chromium not found, will use default Chrome")

# === Generate Roll Numbers (72→99, A1→D9) ===
def generate_roll_numbers():
    rolls = []
    for num in range(72, 100):
        code = str(num)
        if code in ["80", "88"]:
            continue
        rolls.append(BASE_PREFIX + code)
    for letter in ["A", "B", "C", "D"]:
        for d in range(0, 10):
            code = f"{letter}{d}"
            if code == "A0":
                continue
            rolls.append(BASE_PREFIX + code)
    print(f"📋 Generated {len(rolls)} roll numbers")
    return rolls

# === Scraper Worker ===
def process_roll(rollP):
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.set_page_load_timeout(10)
            wait = WebDriverWait(driver, 5)

            driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")
            wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(rollP)
            driver.find_element(By.ID, "btnNext").click()
            wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(rollP)
            driver.find_element(By.ID, "btnSubmit").click()
            wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()

            # Get Overall Attendance
            wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_grdSubject")))
            table = driver.find_element(By.ID, "ctl00_cpStud_grdSubject")
            soup = BeautifulSoup(table.get_attribute('outerHTML'), 'html.parser')
            rows = soup.find_all('tr')[1:]  # Skip header row

            classes_held = []
            attended = []
            percentages = []
            subjects = []
            for row in rows:
                cells = [td.text.strip() for td in row.find_all('td')]
                cells = [cell if cell != '\xa0' else '' for cell in cells]
                subjects.append(cells[1].split(' : ')[0])  # Subject code
                classes_held.append(cells[3])  # Classes Held
                attended.append(cells[4])  # Classes Attended
                percentages.append(cells[5])  # Attendance %

            # Overall Attended and Percentage from Total row (last row)
            overall_attended = attended[-1] if attended else ''
            overall_percentage = percentages[-1] if percentages else ''

            print(f"✅ {rollP} → Overall Attended: {overall_attended}, Percentage: {overall_percentage}")
            return {
                'roll': rollP[:-1],  # Remove P
                'overall_attended': overall_attended,
                'overall_percentage': overall_percentage,
                'classes_held': classes_held,
                'subjects': subjects,
                'attended': attended,
                'percentages': percentages
            }

        except Exception as e:
            print(f"⚠️ Attempt {attempt} failed: {rollP} — {e}")
            time.sleep(0.5)
        finally:
            try:
                driver.quit()
            except:
                pass

    print(f"❌ Max attempts failed: {rollP}")
    return {
        'roll': rollP[:-1], 
        'overall_attended': '', 
        'overall_percentage': '', 
        'classes_held': [], 
        'subjects': [], 
        'attended': [], 
        'percentages': []
    }

# === Prepare new column in Sheet2 and Subject Sheets ===
def prepare_new_column(sheet):
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    current_datetime = ist_time.strftime("%Y-%m-%d %I:%M %p")
    sheet.insert_cols([[]], 3)
    sheet.update_cell(10, 3, current_datetime)
    print(f"📅 Created new column C in {sheet.title} with timestamp: {current_datetime}")
    return 3

# === Get existing roll → row mapping from Sheet1 ===
def get_roll_row_mapping():
    all_rows = sheet1.get_all_values()
    roll_map = {}
    for idx, row in enumerate(all_rows[26:], start=27):  # Start from row 27
        if len(row) > 0 and row[0].strip():
            roll_map[row[0].strip()] = idx
    return roll_map

# === Create or Get Subject Sheets ===
def get_or_create_sheets():
    sheets = {}
    for subject in ["Overall %"] + SUBJECTS:
        try:
            sheets[subject] = spreadsheet.worksheet(subject)
        except gspread.exceptions.WorksheetNotFound:
            sheets[subject] = spreadsheet.add_worksheet(title=subject, rows=100, cols=10)
            # Initialize with roll numbers and names from Sheet1
            sheet1_data = sheet1.get('B27:C91')
            if sheet1_data:
                sheets[subject].update('A11:B' + str(11 + len(sheet1_data) - 1), sheet1_data)
    return sheets

# === Run Scraping and Update All Sheets ===
def run_parallel_scraping():
    rolls = generate_roll_numbers()
    roll_to_row = get_roll_row_mapping()
    print(f"🗂 Found {len(roll_to_row)} rolls mapped in Sheet1")

    # Get or create sheets
    sheets = get_or_create_sheets()
    col_position = prepare_new_column(sheets["Overall %"])
    for subject in SUBJECTS:
        prepare_new_column(sheets[subject])

    rolls_with_P = [r + "P" for r in rolls]
    batch_size = MAX_THREADS
    total_batches = (len(rolls_with_P) + batch_size - 1) // batch_size

    for batch_index in range(total_batches):
        start = batch_index * batch_size
        end = start + batch_size
        batch_rolls = rolls_with_P[start:end]
        print(f"\n🚀 Batch {batch_index + 1}/{total_batches}: {batch_rolls}")

        batch_results = {}
        with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
            futures = {executor.submit(process_roll, roll): roll for roll in batch_rolls}
            for future in as_completed(futures):
                batch_results[future.result()['roll']] = future.result()

        # Update sheets for batch
        for roll, data in batch_results.items():
            if roll not in roll_to_row:
                print(f"⚠️ Roll {roll} not found in Sheet1 → skipped")
                continue

            row_idx = roll_to_row[roll]
            subject_data = {data['subjects'][i]: (data['attended'][i], data['percentages'][i]) 
                           for i in range(len(data['subjects'])-1)}  # Exclude Total

            # Update Sheet1 (D:AE only)
            row_data = [
                data['overall_attended'],  # Column D: Overall Attended
                data['overall_percentage'],  # Column E: Overall Percentage
                subject_data.get('DAA', ('0', ''))[0], subject_data.get('DAA', ('0', ''))[1],
                subject_data.get('CN', ('0', ''))[0], subject_data.get('CN', ('0', ''))[1],
                subject_data.get('DEVOPS', ('0', ''))[0], subject_data.get('DEVOPS', ('0', ''))[1],
                subject_data.get('PPL', ('0', ''))[0], subject_data.get('PPL', ('0', ''))[1],
                subject_data.get('NLP', ('0', ''))[0], subject_data.get('NLP', ('0', ''))[1],
                subject_data.get('CN LAB', ('0', ''))[0], subject_data.get('CN LAB', ('0', ''))[1],
                subject_data.get('DEVOPS LAB', ('0', ''))[0], subject_data.get('DEVOPS LAB', ('0', ''))[1],
                subject_data.get('ACS LAB', ('0', ''))[0], subject_data.get('ACS LAB', ('0', ''))[1],
                subject_data.get('IPR', ('0', ''))[0], subject_data.get('IPR', ('0', ''))[1],
                subject_data.get('Sports', ('0', ''))[0], subject_data.get('Sports', ('0', ''))[1],
                subject_data.get('Men', ('0', ''))[0], subject_data.get('Men', ('0', ''))[1],
                subject_data.get('Assoc', ('0', ''))[0], subject_data.get('Assoc', ('0', ''))[1],
                subject_data.get('Lib', ('0', ''))[0], subject_data.get('Lib', ('0', ''))[1]
            ]
            sheet1.update(f'D{row_idx}:AE{row_idx}', [row_data])

            # Update I8:I21 and C24 only for 237Z1A0575
            if roll == '237Z1A0575':
                sheet1.update('I8:I21', [[val] for val in data['classes_held']])
                sheet1.update('C24', '2025-07-23 21:15')

            # Update Sheet2 (Overall %) and Subject Sheets
            sheets["Overall %"].update_cell(row_idx, col_position, data['overall_percentage'])
            for subject in SUBJECTS:
                percentage = subject_data.get(subject, ('0', ''))[1]
                sheets[subject].update_cell(row_idx, col_position, percentage)

        time.sleep(1)

    print("\n✅ All rolls processed & attendance updated across all sheets!")

# === MAIN ===
if __name__ == "__main__":
    run_parallel_scraping()
