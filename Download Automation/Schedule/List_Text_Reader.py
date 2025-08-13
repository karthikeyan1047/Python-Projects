from selenium import webdriver
import pymsgbox as msg
import _functions as cfx
import xlwings as xw
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import datetime, time, sys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert

url = "https://rcmbi.processmed.ae/login"
user_id = 'rakbi-dhpo@processmed.ae'
user_password = "xL=0UX+:sb=@x#n1X).y"

def wait_for_element(by, value, timeout=20):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))

def wait_for_clickable_element(by, value, timeout=20):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))

def dropdown_select(dropdown, option):
    wait = WebDriverWait(driver, 20)
    dd = wait.until(EC.element_to_be_clickable((By.ID, dropdown)))      # f"//span[contains(text(), '{locator}')]/following::span[@role='combobox']"
    dd.click()
    opt = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, f"//div[contains(@class, 'p-overlay')]//*[normalize-space()='{option}']")
        )
    )
    try:
        opt.click()
        time.sleep(2)
    except:
        driver.execute_script("arguments[0].scrollIntoView(true);", opt)
        time.sleep(0.5)
        opt.click()
        time.sleep(2)

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

driver = webdriver.Chrome(options=options)
driver.implicitly_wait(60)
wait = WebDriverWait(driver, 30)

driver.get(url)
wait_for_element(By.ID, "email").clear()
driver.find_element(By.ID, "email").send_keys(user_id)
driver.find_element(By.ID, "password").clear()
driver.find_element(By.ID, "password").send_keys(user_password)
time.sleep(1)
sign_in_button = wait_for_clickable_element(By.XPATH, "//button[text()='Sign In']", 30)
sign_in_button.click()
time.sleep(2)
driver.execute_script("document.body.style.zoom='90%'")
time.sleep(1)

# dd_id = 'pn_id_7'
# wait_for_clickable_element(By.ID, f'{dd_id}').click()
# elements = driver.find_elements(By.XPATH, f"//*[@id='{dd_id}']/p-overlay//ul/p-selectitem/li")

dropdown_select('pn_id_7', 'internalcomment')

elements = driver.find_elements(By.XPATH, "//div[contains(text(), 'Select Columns')]/following::div[contains(@class, 'source-list-container')]//ul/li")
# elements = driver.find_elements(By.XPATH, "//div[contains(text(), 'Select Columns')]/following::div[contains(@class, 'target-list-container')]//ul/li")

for idx, n in enumerate(elements, start=1):
    print(f"{idx} - {n.text}")
