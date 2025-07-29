from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# Setup headless browser
chrome_options = Options()
chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")

driver = webdriver.Chrome(options=chrome_options)
wait = WebDriverWait(driver, 15)

try:
    # Step 1: Open BeeSERP Login Page
    driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")
    
    # Step 2: Enter Hall Ticket Number and Click Next
    wait.until(EC.presence_of_element_located((By.ID, "txtHTNO"))).send_keys("237Z1A0575P")
    driver.find_element(By.ID, "btnNext").click()

    # Step 3: Enter Password and Click Login
    wait.until(EC.presence_of_element_located((By.ID, "txtPwd"))).send_keys("237Z1A0575P")
    driver.find_element(By.ID, "btnLogin").click()

    # Step 4: Click on Student Dashboard
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Click Here to go Student Dashbord"))).click()

    # Step 5: Wait for Subject-wise Attendance Table
    table_div = wait.until(EC.presence_of_element_located((By.ID, "ctl00_cpStud_PanelSubjectwise")))
    rows = table_div.find_elements(By.XPATH, ".//tr")[1:]  # skip header

    classes_held_list = []
    for row in rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) >= 3:
            held_value = cells[2].text.strip()  # Usually 3rd column = "Classes Held"
            try:
                classes_held_list.append(int(held_value))
            except:
                classes_held_list.append(0)

    print("Extracted Classes Held:", classes_held_list)

finally:
    driver.quit()

# Step 6: Insert into Google Sheet in range D8:D21
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1hu2BoArCojZJGHNODGuaHNE3agS4wylbr_MFOeEWiKI/edit#gid=0").worksheet("Attendence CSE-B(2023-27)")

# Make sure list has exactly 14 values (D8:D21)
while len(classes_held_list) < 14:
    classes_held_list.append(0)
classes_held_list = classes_held_list[:14]

# Convert to 2D array and update range
update_range = 'D8:D21'
data = [[val] for val in classes_held_list]
sheet.update(update_range, data)

print("✅ Classes Held inserted into D8:D21")
