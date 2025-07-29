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
chrome_options.add_argument("--disable-notifications")  # Prevent pop-ups

# Detect chromium-browser path for compatibility (e.g., GitHub Actions)
chrome_path = which("chromium-browser")
if chrome_path:
    print(f"✅ Found Chromium at: {chrome_path}")
    chrome_options.binary_location = chrome_path
else:
    print("⚠️ Chromium not found, using default Chrome")

# Initialize WebDriver
driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 30)  # Increased timeout to 30 seconds

try:
    # Step 1: Open BeeSERP Login Page
    driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")
    print("✅ Opened login page")

    # Step 2: Enter Hall Ticket Number and Click Next
    try:
        username_field = wait.until(EC.presence_of_element_located((By.ID, "txtUserName")))
        username_field.send_keys("237Z1A0575P")
        driver.find_element(By.ID, "btnNext").click()
        print("✅ Entered username and clicked Next")
    except Exception as e:
        print(f"⚠️ Failed to locate txtUserName or btnNext: {e}")

    # Step 3: Enter Password and Click Login
    try:
        password_field = wait.until(EC.presence_of_element_located((By.ID, "txtPassword")))
        password_field.send_keys("237Z1A0575P")
        driver.find_element(By.ID, "btnSubmit").click()
        print("✅ Entered password and clicked Submit")
    except Exception as e:
        print(f"⚠️ Failed to locate txtPassword or btnSubmit: {e}")

    # Step 4: Click on Student Dashboard
    try:
        dashboard_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Click Here to go Student Dashbord")))
        dashboard_link.click()
        print("✅ Clicked on Student Dashboard")
    except Exception as e:
        print(f"⚠️ Failed to locate dashboard link: {e}")

    # Step 5: Wait for Subject-wise Attendance Table
    classes_held_list = []  # Initialize to avoid undefined variable error
    total_held = 0
    try:
        # Wait for the table to be fully loaded with specific data
        wait.until(EC.visibility_of_element_located((By.XPATH, "//table[@id='ctl00_cpStud_grdSubject']//tr[td[contains(text(), 'DEVOPS')]]")))
        time.sleep(5)  # Increased delay for dynamic content
        table_div = wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_PanelSubjectwise")))
        rows = table_div.find_elements(By.XPATH, ".//tr")[1:]  # Skip header

        for i, row in enumerate(rows, 1):
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) >= 6:  # Ensure row has enough columns
                subject_name = cells[1].text.strip()  # Subject is in 2nd column
                held_value = cells[3].text.strip()  # Classes Held is 4th column (index 3)
                print(f"Row {i}: Subject: {subject_name}, Classes Held: {held_value}")  # Detailed debug
                try:
                    value = int(held_value)
                    if subject_name != "-":  # Exclude "Total" from the main list
                        classes_held_list.append(value)
                    else:
                        total_held = value  # Capture the total value
                except ValueError:
                    if subject_name != "-":
                        classes_held_list.append(0)

        print("✅ Extracted Classes Held:", classes_held_list)
        print(f"✅ Total Classes Held: {total_held}")
        if len(classes_held_list) != 13:
            print(f"⚠️ Warning: Expected 13 rows, got {len(classes_held_list)}")
    except Exception as e:
        print(f"⚠️ Failed to locate or process attendance table: {e}")
        print("ℹ️ Using default empty list due to table loading failure")

    # Save page source for debugging
    with open("page_source.html", "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    print("✅ Saved page source to page_source.html")

finally:
    driver.quit()

# Step 6: Insert into Google Sheet in range D8:D21
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

try:
    # Open sheet by ID
    sheet = client.open_by_key("1hu2BoArCojZJGHNODGuaHNE3agS4wylbr_MFOeEWiKI").worksheet("Attendence CSE-B(2023-27)")

    # Ensure list has exactly 14 values (D8:D21)
    while len(classes_held_list) < 14:
        classes_held_list.append(0)
    classes_held_list = classes_held_list[:14]

    # Convert to 2D array and update range
    update_range = 'D8:D21'
    data = [[val] for val in classes_held_list]
    sheet.update(values=data, range_name=update_range)

    print("✅ Classes Held inserted into D8:D21")
except Exception as e:
    print(f"⚠️ Failed to update Google Sheet: {e}")
