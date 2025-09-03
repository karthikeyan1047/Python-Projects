from selenium import webdriver
import pyautogui as gui
import _functions as cfx
import time, clipboard, os
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
import pandas as pd

empty_data = pd.DataFrame()
dwl_folder = r"C:\Users\karthikeyans\Downloads"

def rec_date_range(datefrm, dateto):
    def pick_date(date):
        wait = WebDriverWait(driver, 3)
        day, day_num = date.strftime('%d'), date.day
        month, month_num = date.strftime('%b'), (date.month - 1)
        year, year_num = date.strftime('%Y'), date.year
        date_label = f"{year_num}-{month_num}-{day_num}"
        
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "p-datepicker-select-year"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[text()=' {year} ']"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[text()=' {month} ']"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[@data-date='{date_label}']"))).click()

    date_input = driver.find_element(By.CSS_SELECTOR, "p-datepicker input")
    date_input.click()

    pick_date(datefrm)
    pick_date(dateto)

def wait_for_element(by, value, timeout=20):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))

def wait_for_clickable_element(by, value, timeout=20):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))

def dropdown_select(dropdown, option):
    wait = WebDriverWait(driver, 20)
    dd = wait.until(EC.element_to_be_clickable((By.ID, dropdown)))
    dd.click()
    opt = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, f"//div[contains(@class, 'p-overlay')]//*[normalize-space()='{option}']")
        )
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", opt)
    opt.click()
    time.sleep(1)

prompt_text = (
    "1. rakbi-dhpo@processmed.ae\n"
    "2. rakbi-rypd@processmed.ae\n"
    "3. rakgroupbi-dhpo@processmed.ae\n"
    "4. mmlbi-rypd@processmed.ae\n"
    "5. simcbi-dhpo@processmed.ae\n"
    "6. simcbi-rypd@processmed.ae\n"
    "7. mdsbi-dhpo@processmed.ae"
)
choice = int(cfx.inputbox(title="Username Select", prompt=prompt_text))

mapping = {
    1 : 'rakbi-dhpo@processmed.ae',
    2 : 'rakbi-rypd@processmed.ae',
    3 : 'rakgroupbi-dhpo@processmed.ae',
    4 : 'mmlbi-rypd@processmed.ae',
    5 : 'simcbi-dhpo@processmed.ae',
    6 : 'simcbi-rypd@processmed.ae',
    7 : 'mdsbi-dhpo@processmed.ae'
}

user_email = mapping.get(choice, "Invalid choice")

url = "https://rcmbi.processmed.ae/login"
template = "Resubmission_Audit"
dateparameter = "lastTransactionDate"
usernames_dict = {
    'rakbi-dhpo@processmed.ae' : [('RAK', 'BI_DHPO_IPOP'), ('RAKP', 'BI_DHPO_PH')],
    'rakbi-rypd@processmed.ae' : [('RMCH', 'BI_RPO_L7026'), ('RMCG', 'BI_RPO_L5223'), ('RMCJ', 'BI_RPO_L5136'), ('RAK', 'BI_RPO_IPOP'), ('RMCP', 'BI_RPO_L206'), ('RAKP', 'BI_RPO_PH')], 
    'rakgroupbi-dhpo@processmed.ae' : [('SMCL', 'XML_BI_SMCL'), ('RMCH', 'XML_BI_HOC'), ('RMCG', 'XML_BI_GOC'), ('RMCJ', 'XML_BI_JOC'), ('MMLN', 'XML_BI_MML')],
    'mmlbi-rypd@processmed.ae' : [('MML', 'BI_RPO_L5741')],
    'simcbi-dhpo@processmed.ae' : [('SIMC', 'XML_BI_443')],
    'simcbi-rypd@processmed.ae' : [('SIMC', 'BI_RPO_L5887')],
    'mdsbi-dhpo@processmed.ae' : [('MDS', 'XML_BI_med')]
}
passwords_dict = {
    'rakbi-dhpo@processmed.ae' : "xL=0UX+:sb=@x#n1X).y",
    'rakbi-rypd@processmed.ae' : "kQm?f_8U1uV7Bf>u.By}",
    'rakgroupbi-dhpo@processmed.ae' : "xL=0UX+:sb=@x#n1X).y",
    'mmlbi-rypd@processmed.ae' : "NU]-k2giuwkF?jWB4dnL",
    'simcbi-dhpo@processmed.ae' : "2+o3M0iFxvVNo?2n)Fyi",
    'simcbi-rypd@processmed.ae' : "}:wmJ.9B.6-b4-i^^PTM",
    'mdsbi-dhpo@processmed.ae' : "aN.FhD1KPjDDtt-wq7:J"
}

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
options.add_experimental_option("detach", True)
options.add_experimental_option("prefs", prefs)

zoom = '90%'

driver = webdriver.Chrome(options=options)
driver.implicitly_wait(60)

wait = WebDriverWait(driver, 20)

username = user_email
password = passwords_dict[user_email]
datefrm = cfx.get_date("Start Date : ")
dateto = cfx.get_date("End Date : ")

driver.get(url)
wait_for_element(By.ID, "email").clear()
driver.find_element(By.ID, "email").send_keys(username)
driver.find_element(By.ID, "password").clear()
driver.find_element(By.ID, "password").send_keys(password)
time.sleep(1)
sign_in_button = wait_for_clickable_element(By.XPATH, "//button[text()='Sign In']", 30)
sign_in_button.click()
time.sleep(2)
driver.execute_script(f"document.body.style.zoom='{zoom}'")
time.sleep(1)
wait_for_clickable_element(By.XPATH, "//i[contains(@class, 'pi-bars')]").click()
time.sleep(0.5)
wait_for_clickable_element(By.XPATH, "//i[contains(@class, 'pi-sun')]").click()
time.sleep(0.5)

# # TEMPLATE-SELECT
dropdown_select('pn_id_7', 'Resubmission_Audit')

# # DATE-COLUMN-SELECT
dropdown_select('pn_id_17', 'lastTransactionDate')

# # DATE-PICKER-DATE-SELECT
rec_date_range(datefrm, dateto)

centers = usernames_dict[user_email]
for (center, flname) in centers:
    dropdown_select("pn_id_11", center)
    time.sleep(1)

    schedule_button = wait_for_clickable_element(By.XPATH, "//button[.//span[text()='Schedule']]", 30)                                                    
    schedule_button.click()
    WebDriverWait(driver, 300).until(EC.alert_is_present())
    time.sleep(1.5)

    alert = Alert(driver)
    alert_text0 = alert.text
    alert_text = alert_text0.encode('ascii', errors='replace').decode('ascii')
    if 'No results found' in alert_text or 'Please contact administrator' in alert_text:
        no_file_path = os.path.join(dwl_folder, f"{flname}.xlsx")
        empty_data.to_excel(no_file_path, sheet_name="Report", index=False)

    Alert(driver).accept()
    time.sleep(1.5)

report_button = wait_for_clickable_element(By.XPATH, "//button[.//span[text()='Reports']]", 30)
report_button.click()

for (center, flname) in centers:

    dropdown_select('clientName', center)

    refresh_button = wait_for_clickable_element(By.XPATH, "//button[.//span[text()='Refresh']]", 30)
    refresh_button.click()

    center_find = wait_for_element(By.XPATH, f"//table//tr/td[2][text()=' {center} ']", 30)

    time.sleep(1)
    dwl_button = wait_for_element(By.XPATH, "//*[@id='pn_id_1-table']/tbody/tr[1]/td[9]/p-button/button", 30)
    dwl_button.click()
    
    time.sleep(1)
    clipboard.copy(flname)
    time.sleep(1)
    gui.hotkey('ctrl', 'v')
    time.sleep(1)
    gui.hotkey('enter')
    time.sleep(1)
    gui.hotkey('y')


driver.quit()