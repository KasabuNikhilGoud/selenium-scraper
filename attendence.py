from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
from concurrent.futures import ThreadPoolExecutor, as_completed
from shutil import which
from bs4 import BeautifulSoup
import gspread
import time
from datetime import datetime
from zoneinfo import ZoneInfo

# === CONFIG ===
SHEET_ID = "1Rk3eNqEhbuDIgu3Zx4_CwOZCnFlLm6Vr9obVzYl_zr4"
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

# === Generate Roll Numbers ===
def generate_roll_numbers():
    rolls = []
    for num in range(72, 100):
        if str(num) in ["80", "88"]:
            continue
        rolls.append(BASE_PREFIX + str(num))
    for letter in ["A", "B", "C", "D"]:
        for d in range(1, 10):
            rolls.append(BASE_PREFIX + f"{letter}{d}")
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

            wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_grdSubject")))
            table = driver.find_element(By.ID, "ctl00_cpStud_grdSubject")
            soup = BeautifulSoup(table.get_attribute('outerHTML'), 'html.parser')
            rows = soup.find_all('tr')[1:]

            classes_held, attended, percentages, subjects = [], [], [], []
            for row in rows:
                cells = [td.text.strip().replace('\xa0', '') for td in row.find_all('td')]
                subjects.append(cells[1].split(' : ')[0])
                classes_held.append(cells[3])
                attended.append(cells[4])
                percentages.append(cells[5])

            return {
                'roll': rollP[:-1],
                'overall_attended': attended[-1],
                'overall_percentage': percentages[-1],
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
    return {'roll': rollP[:-1], 'overall_attended': '', 'overall_percentage': '', 'classes_held': [], 'subjects': [], 'attended': [], 'percentages': []}

# === Get Roll Mapping ===
def get_roll_row_mapping():
    roll_map = {}
    for i, row in enumerate(sheet1.get_all_values()[26:], start=27):
        if row and row[0].strip():
            roll_map[row[0].strip()] = i
    return roll_map

# === Get or Create Subject Sheets ===
def get_or_create_sheets():
    sheets = {}
    for subject in ["Overall %"] + SUBJECTS:
        try:
            sheets[subject] = spreadsheet.worksheet(subject)
        except gspread.exceptions.WorksheetNotFound:
            sheets[subject] = spreadsheet.add_worksheet(title=subject, rows=100, cols=10)
            sheet1_data = sheet1.get('B27:C91')
            if sheet1_data:
                sheets[subject].update('A11:B' + str(11 + len(sheet1_data) - 1), sheet1_data)
    return sheets

# === Prepare Column ===
def prepare_new_column(sheet):
    ist_time = datetime.now(ZoneInfo("Asia/Kolkata"))
    current_datetime = ist_time.strftime("%Y-%m-%d %I:%M %p")
    sheet.insert_cols([[]], 3)
    sheet.update_cell(10, 3, current_datetime)
    return 3

# === Main Runner ===
def run_parallel_scraping():
    rolls = generate_roll_numbers()
    roll_map = get_roll_row_mapping()
    sheets = get_or_create_sheets()
    col_pos = prepare_new_column(sheets["Overall %"])
    for subject in SUBJECTS:
        prepare_new_column(sheets[subject])

    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        futures = {executor.submit(process_roll, roll + "P"): roll for roll in rolls}
        for future in as_completed(futures):
            data = future.result()
            roll = data['roll']
            if roll not in roll_map:
                continue
            row_idx = roll_map[roll]
            sheet1.update(f'I8:I21', [[v] for v in data['classes_held']] if roll == '237Z1A0575' else [])
            sheet1.update('C24', datetime.now().strftime("%Y-%m-%d %H:%M") if roll == '237Z1A0575' else '')

            subj_data = {data['subjects'][i]: (data['attended'][i], data['percentages'][i]) for i in range(len(data['subjects'])-1)}
            row_data = [
                data['overall_attended'], data['overall_percentage'],
                *[val for subj in SUBJECTS for val in subj_data.get(subj, ('0', ''))]
            ]
            sheet1.update(f'D{row_idx}:AE{row_idx}', [row_data])

            sheets["Overall %"].update_cell(row_idx, col_pos, data['overall_percentage'])
            for subject in SUBJECTS:
                sheets[subject].update_cell(row_idx, col_pos, subj_data.get(subject, ('0', ''))[1])
    print("\n✅ All rolls processed & attendance updated!")

# === MAIN ===
if __name__ == "__main__":
    run_parallel_scraping()
