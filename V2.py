from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import time
from shutil import which

# Setup Chrome options
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--log-level=3")
chrome_options.add_argument("--window-size=1280,800")

# Detect chromium-browser path for compatibility (e.g., GitHub Actions)
chrome_path = which("chromium-browser")
if chrome_path:
    print(f"✅ Found Chromium at: {chrome_path}")
    chrome_options.binary_location = chrome_path
else:
    print("⚠️ Chromium not found, using default Chrome")

# Initialize WebDriver
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 10)  # Reduced timeout for faster failure

try:
    # Step 1: Open BeeSERP Login Page
    driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")

    # Step 2: Enter Hall Ticket Number and Click Next
    try:
        username_field = wait.until(EC.presence_of_element_located((By.ID, "txtUserName")))
        username_field.send_keys("237Z1A0575P")
        driver.find_element(By.ID, "btnNext").click()
    except:
        print("⚠️ Failed to locate txtUserName or btnNext, check element IDs")

    # Step 3: Enter Password and Click Login
    try:
        password_field = wait.until(EC.presence_of_element_located((By.ID, "txtPassword")))
        password_field.send_keys("237Z1A0575P")
        driver.find_element(By.ID, "btnSubmit").click()
    except:
        print("⚠️ Failed to locate txtPassword or btnSubmit, check element IDs")

    # Step 4: Click on Student Dashboard
    try:
        dashboard_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Click Here to go Student Dashbord")))
        dashboard_link.click()
    except:
        print("⚠️ Failed to locate dashboard link, check link text")

    # Step 5: Wait for Subject-wise Attendance Table
    try:
        table_div = wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_PanelSubjectwise")))
        rows = table_div.find_elements(By.XPATH, ".//tr")[1:]  # Skip header

        classes_held_list = []
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 3:
                held_value = cells[2].text.strip()  # 3rd column = "Classes Held"
                try:
                    classes_held_list.append(int(held_value))
                except:
                    classes_held_list.append(0)

        print("✅ Extracted Classes Held:", classes_held_list)
    except:
        print("⚠️ Failed to locate or process attendance table")

finally:
    driver.quit()

# Step 6: Insert into Google Sheet in range D8:D21
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

try:
    # Open sheet by ID
    sheet = client.open_by_key("1Rk3eNqEhbuDIgu3Zx4_CwOZCnFlLm6Vr9obVzYl_zr4").worksheet("Attendence CSE-B(2023-27)")

    # Ensure list has exactly 14 values (D8:D21)
    while len(classes_held_list) < 14:
        classes_held_list.append(0)
    classes_held_list = classes_held_list[:14]

    # Convert to 2D array and update range
    update_range = 'D8:D21'
    data = [[val] for val in classes_held_list]
    sheet.update(update_range, data)

    print("✅ Classes Held inserted into D8:D21")
except Exception as e:
    print(f"⚠️ Failed to update Google Sheet: {e}")
