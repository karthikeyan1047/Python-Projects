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
import pyautogui as gui

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

url = "https://rcmbi.processmed.ae/login"
username, password = 'rakbi-dhpo@processmed.ae', "xL=0UX+:sb=@x#n1X).y"
driver = webdriver.Chrome(options=options)
driver.implicitly_wait(60)
wait = WebDriverWait(driver, 20)

login(username, password)

wait_for_clickable_element(By.XPATH, "//button[.//span[text()='Reports']]", 30).click()

center = 'RAK'
dropdown_select('clientName', center)
wait_for_clickable_element(By.XPATH, "//button[.//span[text()='Refresh']]", 30).click()

old_names = [""]
old_names = sorted(old_names)
check_pg_for = old_names[0]
a, pg = True, 0
while a:
    wait.until(
        lambda d: len(d.find_elements(By.XPATH, "//*[@id='pn_id_1-table']/tbody/tr")) >= 10
    )
    rows = driver.find_elements(By.XPATH, "//*[@id='pn_id_1-table']/tbody/tr")
    for r in range(1, len(rows)+1):
        sch_filename = driver.find_element(By.XPATH, f"//*[@id='pn_id_1-table']/tbody/tr[{r}]/td[5]").text
        if sch_filename == check_pg_for:
            a = False
            break

    pg+=1
    if a:
        wait_for_clickable_element(By.XPATH, "//button[@aria-label='Next Page']").click()

wait_for_clickable_element(By.XPATH, "//button[@aria-label='First Page']").click()

for i in range(1, pg+1):
    wait.until(
        lambda d: len(d.find_elements(By.XPATH, "//*[@id='pn_id_1-table']/tbody/tr")) >= 10
    )
    rows = driver.find_elements(By.XPATH, "//*[@id='pn_id_1-table']/tbody/tr")
    for r in range(1, len(rows)+1):
        sch_filename = driver.find_element(By.XPATH, f"//*[@id='pn_id_1-table']/tbody/tr[{r}]/td[5]").text
        if sch_filename in old_names:
            dwl_button = wait_for_clickable_element(By.XPATH, f"//*[@id='pn_id_1-table']/tbody/tr[{r}]/td[9]/p-button/button").click()
            time.sleep(1)
            gui.hotkey('enter')
    wait_for_clickable_element(By.XPATH, "//button[@aria-label='Next Page']").click()


