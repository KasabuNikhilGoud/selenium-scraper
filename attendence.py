from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from shutil import which
import time

# === CONFIG ===
ROLL_P = "237Z1A0575P"
ROLL = ROLL_P[:-1]
SHEET_ID = "1Rk3eNqEhbuDIgu3Zx4_CwOZCnFlLm6Vr9obVzYl_zr4"

# === Google Sheets Setup ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

# === Chrome Setup ===
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument("--window-size=1280,800")

# For GitHub Actions / Linux
chrome_path = which("chromium-browser")
if chrome_path:
    chrome_options.binary_location = chrome_path

# === Main Function ===
def test_single_student_subject_attendance():
    try:
        driver = webdriver.Chrome(options=chrome_options)
        driver.set_page_load_timeout(10)
        wait = WebDriverWait(driver, 5)

        # Step 1: Open login page
        driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")

        # Step 2: Enter roll and go to password
        wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(ROLL_P)
        driver.find_element(By.ID, "btnNext").click()

        # Step 3: Enter password and login
        wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(ROLL_P)
        driver.find_element(By.ID, "btnSubmit").click()

        # Step 4: Click to go to Dashboard
        wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()

        # Step 5: Wait for subject table
        wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_grdSubject")))
        table = driver.find_element(By.ID, "ctl00_cpStud_grdSubject")
        rows = table.find_elements(By.TAG_NAME, "tr")[1:]  # Skip header

        subject_data = []
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) >= 6:
                slno = cols[0].text.strip()
                subject = cols[1].text.strip()
                faculty = cols[2].text.strip()
                held = cols[3].text.strip()
                attended = cols[4].text.strip()
                percent = cols[5].text.strip()
                subject_data.append([slno, subject, faculty, held, attended, percent])

        print(f"✅ Found {len(subject_data)} subjects for {ROLL}")

        # Step 6: Insert into new sheet named "237Z1A0575"
        try:
            target_sheet = client.open_by_key(SHEET_ID).worksheet(ROLL)
        except gspread.exceptions.WorksheetNotFound:
            target_sheet = client.open_by_key(SHEET_ID).add_worksheet(title=ROLL, rows="100", cols="10")

        target_sheet.clear()
        target_sheet.append_row(["SlNo", "Subject", "Faculty", "Classes Held", "Classes Attended", "Att %"])
        for row in subject_data:
            target_sheet.append_row(row)

        print(f"📥 Subject-wise data inserted into sheet: {ROLL}")

    except Exception as e:
        print(f"❌ Error: {e}")
    finally:
        try:
            driver.quit()
        except:
            pass

# === Run Main ===
if __name__ == "__main__":
    test_single_student_subject_attendance()
