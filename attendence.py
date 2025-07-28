import time
import json
import re
import os
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ✅ Handle bs4 without installing globally
try:
    from bs4 import BeautifulSoup
except ImportError:
    import subprocess
    import sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "beautifulsoup4"])
    from bs4 import BeautifulSoup

# Google Sheets setup
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

SPREADSHEET_ID = "1hu2BoArCojZJGHNODGuaHNE3agS4wylbr_MFOeEWiKI"
sheet = client.open_by_key(SPREADSHEET_ID)

main_sheet = sheet.worksheet("Attendence CSE-B(2023-27)")
all_sheets = sheet.worksheets()

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--no-sandbox")

# ✅ Read only I8:I21 range (specific 14 students)
roll_cells = main_sheet.range("I8:I21")
rolls = [(cell.row, cell.value.strip()) for cell in roll_cells if cell.value.strip()]

for row_index, roll in rolls:
    driver = webdriver.Chrome(options=chrome_options)
    try:
        driver.get("https://exams-nnrg.in/BeeSERP/Login.aspx")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "txtUserId"))
        ).send_keys(roll)

        driver.find_element(By.ID, "btnNext").click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "txtPwd"))
        ).send_keys(roll)

        driver.find_element(By.ID, "btnLogin").click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.LINK_TEXT, "Click Here to go Student Dashbord"))
        ).click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_cpStud_lblTotalPercentage"))
        )

        page_source = driver.page_source
        soup = BeautifulSoup(page_source, "html.parser")

        subject_blocks = soup.find_all("table", {"class": "table table-bordered table-hover table-striped"})

        subject_data = {}
        for table in subject_blocks:
            headers = [th.get_text(strip=True) for th in table.find_all("th")]
            if "Subject" not in headers:
                continue
            for row in table.find_all("tr")[1:]:
                cells = row.find_all("td")
                if len(cells) >= 5:
                    subject = cells[0].get_text(strip=True)
                    attended = cells[2].get_text(strip=True)
                    percent = cells[4].get_text(strip=True).replace("%", "")
                    subject_data[subject.upper()] = (attended, percent)

        subject_list = [
            "DAA", "CN", "DEVOPS", "PPL", "NLP",
            "CN LAB", "DEVOPS LAB", "ACS LAB", "IPR",
            "SPORTS", "MENTORING", "ASSOCIATION", "LIBRARY"
        ]

        # ✅ Update Overall % column in main sheet
        main_sheet.update(f"D{row_index}", [[subject_data.get(subj, ("", ""))[1] for subj in subject_list]])

        # ✅ Update each subject sheet
        for ws in all_sheets:
            title = ws.title.strip().upper()
            if title in subject_list:
                col_labels = ws.row_values(25)
                if "Percentage%" in col_labels:
                    percent_col = col_labels.index("Percentage%") + 4
                    attended_col = percent_col - 1
                    ws.update_cell(row_index, attended_col, subject_data.get(title, ("", ""))[0])
                    ws.update_cell(row_index, percent_col, subject_data.get(title, ("", ""))[1])

    except Exception as e:
        print(f"❌ Error for {roll}: {e}")
    finally:
        driver.quit()
