from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from datetime import datetime, timedelta
import time
import pandas as pd
import _functions as cfx
from openpyxl import load_workbook
import xlwings as xw

options = Options()
prefs = {
    "download.prompt_for_download": True,
    "download.default_directory": "",
    "profile.default_content_settings.popups": 1
}
options.add_argument("--start-maximized")
options.add_argument("--disable-popup-blocking")
options.add_argument('--ignore-ssl-errors')
options.add_argument('--ignore-certificate-errors')
options.add_experimental_option("detach", True)     # it will not close the browser automatically
options.add_experimental_option("prefs", prefs)

zoom = '90%'

def wait_for_element(by, value, timeout=20):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))

def wait_for_clickable_element(by, value, timeout=20):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))

def login(username, password):
    driver.get(url)
    wait_for_element(By.ID, "email").clear()
    driver.find_element(By.ID, "email").send_keys(username)
    driver.find_element(By.ID, "password").clear()
    driver.find_element(By.ID, "password").send_keys(password)
    time.sleep(1)
    wait_for_clickable_element(By.XPATH, "//button[text()='Sign In']").click()
    time.sleep(1.5)
    driver.execute_script(f"document.body.style.zoom='{zoom}'")
    time.sleep(0.5)
    wait_for_clickable_element(By.XPATH, "//i[contains(@class, 'pi-bars')]").click()
    time.sleep(0.5)
    wait_for_clickable_element(By.XPATH, "//i[contains(@class, 'pi-sun')]").click()
    time.sleep(0.5)

def dropdown_select(dropdown, option):
    dd = wait_for_clickable_element(By.ID, dropdown)
    dd.click()
    opt = wait_for_clickable_element(
        By.XPATH, f"//div[contains(@class, 'p-overlay')]//*[normalize-space()='{option}']"
    )
    # try:
    #     opt.click()
    # except:
    driver.execute_script("arguments[0].scrollIntoView(true);", opt)
        # time.sleep(0.5)
    opt.click()
    time.sleep(1)

def login_and_select_date_field(username, password, date_field):
    login(username, password)
    dropdown_select('pn_id_17', date_field)

def generate_month_ranges(year, step):
    current_year = datetime.today().year
    current_month = datetime.today().month
    ranges = []

    for i in range(1, 13, step):
        start = i
        end = min(i + step - 1, 12)

        if year < current_year:
            ranges.append((start, end))
        elif year == current_year:
            if start > current_month:
                break
            ranges.append((start, min(end, current_month)))
    
    return ranges

def get_date_range(year, start_month, end_month):
    date_from = datetime(year, start_month, 1)
    if end_month == 12:
        date_to = datetime(year, 12, 31)
    else:
        date_to = datetime(year, end_month + 1, 1) - timedelta(days=1)
    return date_from, date_to

def date_picker(datefrm, dateto):
    def pick_date(date):
        wait = WebDriverWait(driver, 5)
        day = date.day
        month_name, month = date.strftime('%b'), date.month - 1
        year = date.year
        date_label = f"{year}-{month}-{day}"

        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "p-datepicker-select-year"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[text()=' {year} ']"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[text()=' {month_name} ']"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[@data-date='{date_label}']"))).click()

    date_input = driver.find_element(By.CSS_SELECTOR, "p-datepicker input")
    date_input.click()

    pick_date(datefrm)
    pick_date(dateto)

def schedule():
    global row
    schedule_button = wait_for_clickable_element(By.XPATH, "//button[.//span[text()='Schedule']]")
    schedule_button.click()

    WebDriverWait(driver, 300).until(EC.alert_is_present())
    time.sleep(1.5)

    alert = Alert(driver)
    alert_text0 = alert.text
    alert_text = alert_text0.encode('ascii', errors='replace').decode('ascii')
    parts = alert_text.replace(",", "").split(" ")
    file_name, status = parts[0], parts[3]

    print(file_name, status)

    Alert(driver).accept()
    time.sleep(1.5)

file_dict = {
    "Unsettled_Sant" : "HIS",
    "Remittance" : "Sub",
    "Resub_Remittance" : "Resub",
    "RAK" : "H",
    "RAKP" : "P"
}



driver = webdriver.Chrome(options=options)

url = "https://rcmbi.processmed.ae/login"
date_field = 'Encounter.End'
yr_curr = datetime.today().year
email, password = "rakbi-dhpo@processmed.ae", "xL=0UX+:sb=@x#n1X).y"
template = "Resubmission_Audit"
center = "RAK"


login_and_select_date_field(email, password, date_field)

dropdown_select('pn_id_7', template)
dropdown_select('pn_id_11', center)

date_from, date_to = datetime(2025,7,1).date(), datetime(2025, 7, 5)
date_picker(date_from, date_to)
time.sleep(1)

schedule()