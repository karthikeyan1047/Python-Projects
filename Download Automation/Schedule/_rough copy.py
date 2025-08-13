from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

options = Options()
options.add_argument("--start-maximized")
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(options=options)
driver.implicitly_wait(60)
wait = WebDriverWait(driver, 60)

url = "https://www.w3schools.com"
driver.get(url)

search_box = wait.until(EC.visibility_of_element_located((By.ID, "search2")))
search_box.send_keys("python")
time.sleep(1)
search_box.send_keys(Keys.RETURN)
time.sleep(1)

search_box1 = wait.until(EC.visibility_of_element_located((By.ID, "tnb-google-search-input")))
search_box1.send_keys("javascript")
time.sleep(1)
search_box1.send_keys(Keys.RETURN)
time.sleep(1)



