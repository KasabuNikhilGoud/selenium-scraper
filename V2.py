from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# --- Google Sheets Setup ---
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

# Open the sheet
spreadsheet = client.open_by_key("1Rk3eNqEhbuDIgu3Zx4_CwOZCnFlLm6Vr9obVzYl_zr4")
worksheet = spreadsheet.worksheet("Attendence CSE-B(2023-27)")

# --- Selenium Setup ---
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 10)

try:
    # Step 1: Open login page
    driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")

    # Step 2: Enter Hall Ticket Number (username) and go next
    roll_number = "237Z1A0575P"
    driver.find_element(By.ID, "txtHTNO").send_keys(roll_number)
    driver.find_element(By.ID, "btnNext").click()

    # Step 3: Enter password (same as phone number or HTNO if that's how it works)
    wait.until(EC.presence_of_element_located((By.ID, "txtPassword"))).send_keys(roll_number)
    driver.find_element(By.ID, "btnLogin").click()

    # Step 4: Click on "Click Here to go Student Dashbord"
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()

    # Step 5: Wait for Subject-wise attendance to load
    wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_PanelSubjectwise")))

    # Step 6: Get all 'Classes Held' values
    classes_held_elements = driver.find_elements(By.XPATH, '//table[@id="ctl00_cpStud_GridSubjectwise"]/tbody/tr/td[6]')
    classes_held = [el.text for el in classes_held_elements if el.text.strip().isdigit()]

    # Debug print
    print("Classes Held per subject:", classes_held)

    # Step 7: Insert into Google Sheet D8:D21
    update_range = 'D8:D21'
    values = [[val] for val in classes_held]
    worksheet.update(update_range, values)
    print("✅ Successfully inserted classes held into Google Sheet.")

except Exception as e:
    print("❌ Error:", e)

finally:
    driver.quit()
