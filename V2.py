from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import time

# --- Setup Chrome Options ---
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920x1080")
chrome_options.add_argument("--log-level=3")

# --- Setup Google Sheets ---
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
SHEET_ID = "1Rk3eNqEhbuDIgu3Zx4_CwOZCnFlLm6Vr9obVzYl_zr4"  # Your Sheet ID
main_sheet = client.open_by_key(SHEET_ID).worksheet("Attendence CSE-B(2023-27)")

# --- Function to Extract and Update Classes Held ---
def update_classes_held_to_main_sheet():
    rollP = "237Z1A0575P"
    roll = rollP[:-1]  # Password = Roll number without last character

    try:
        driver = webdriver.Chrome(options=chrome_options)
        driver.set_page_load_timeout(10)
        wait = WebDriverWait(driver, 5)

        # Step 1: Login
        driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")
        wait.until(EC.presence_of_element_located((By.ID, "txtUserName"))).send_keys(rollP)
        driver.find_element(By.ID, "btnNext").click()
        wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(rollP)
        driver.find_element(By.ID, "btnSubmit").click()

        # Step 2: Go to Dashboard
        wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()

        # Step 3: Wait for Subject-wise Attendance Table
        wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_grdSubject")))
        table = driver.find_element(By.ID, "ctl00_cpStud_grdSubject")
        rows = table.find_elements(By.TAG_NAME, "tr")[1:]  # skip header row

        classes_held = []
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) >= 4:
                held = cols[3].text.strip()
                classes_held.append(held if held else "0")

        # Step 4: Write to Google Sheet (D8:D21)
        cell_range = f"D8:D{7 + len(classes_held)}"
        cell_list = main_sheet.range(cell_range)

        for i, cell in enumerate(cell_list):
            cell.value = classes_held[i]

        main_sheet.update_cells(cell_list)
        print(f"✅ Inserted {len(classes_held)} Classes Held values to D8:D21 successfully.")

    except Exception as e:
        print(f"❌ Error: {e}")
    finally:
        try:
            driver.quit()
        except:
            pass

# --- Run ---
if __name__ == "__main__":
    update_classes_held_to_main_sheet()
