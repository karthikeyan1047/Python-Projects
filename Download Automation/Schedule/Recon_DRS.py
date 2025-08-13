from selenium import webdriver
import _functions as cfx
import sys, time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from datetime import timedelta

url1 = "https://rcm.processmed.ae:5073/Account/Login?p=rakhospital"
url2 = "https://rcm.processmed.ae:5091/"


usernames1 = ["L1145", "L5067", "L206"]
passwords1 = ["b9W%68Ry&&BK", "bT2&sj9$2.5cL3Sz", "2v<>sW4NzJIRP3A)"]

usernames2 = ["L5136", "L5887", "L7026", "L5741", "L5223", "L153"]
passwords2 = ["C)cC1@8+3K\"X;#p-", "]k,?F1|W:rr4g>1-", "F_1IWb@Pa9a9_$ti", "`i7~06_eY78jA;iR", "|Hj\\}KT210=Mv`28", "X^g285U@(+oalbSJ"]

usernames3 = ['rakhpharmacy', 'rakhospital', 'rakmc-ghalilla', 'smclsharjah', 'SIMCNEW', 'Medstar2015', 'alhamra-rmc', 'rmcjazeera', 'Smcl']
passwords3 = ["4>o14}6)z71Bx4\\a", "lU\"i(.X77+2<3W@p", "Zi92#kW\"Q2`$AOp4", ",5s8QjYb,q}bH6b3", "wQ73r8{rCR>SF@l{", "/*OL~h0z77DF$g1A", "|}y629688@5P^WTf", "3?xq0:Rc'Jjl4>i,", "ym7Yg5aC[VE*0w;5"]


choice = int(cfx.inputbox(title='Choose', prompt='Choose : \n   1. DHPO\n   2. RPO\n   3. RPO - 1\n   4. RPO - 2\n   5. All\n\nEnter Choice : '))

datefrm = cfx.get_date("From Date : ")
dateto = cfx.get_date("To Date : ")
mdate = dateto - timedelta(days=3)
dates = [dateto, mdate, datefrm]
datefrm1 = datefrm.strftime("%m/%d/%Y")
dateto1 = dateto.strftime("%m/%d/%Y")
print(datefrm1)
print(dateto1)

options = Options()
prefs = {
    "download.prompt_for_download": True,
    "download.default_directory": "",
    "profile.default_content_settings.popups": 1
}
options.add_experimental_option("prefs", prefs)
options.add_argument("--start-maximized")
options.add_argument('--ignore-ssl-errors')
options.add_argument('--ignore-certificate-errors')
options.add_argument("--disable-popup-blocking")
options.add_experimental_option('detach', True)


def login_and_filling(username, password, dtfrm, dtto, is_rpo):
        driver = webdriver.Chrome(options=options)
        driver.implicitly_wait(15)
        if is_rpo:
            driver.get(url2)
        else:
            driver.get(url1)
        
        driver.find_element(By.ID, "username").clear()
        driver.find_element(By.ID, "username").send_keys(username)
        driver.find_element(By.ID, "password").clear()
        driver.find_element(By.ID, "password").send_keys(password)
        driver.find_element(By.XPATH, "//*[@id='login-form']/footer/a").click()
        driver.find_element(By.XPATH, "//*[@id='frmdate']").clear()
        driver.find_element(By.XPATH, "//*[@id='frmdate']").send_keys(dtfrm)
        driver.find_element(By.XPATH, "//*[@id='todate']").clear()
        driver.find_element(By.XPATH, "//*[@id='todate']").send_keys(dtto)

        if is_rpo:
            Select(driver.find_element(By.ID, "seldownload")).select_by_visible_text('All')
        
        Select(driver.find_element(By.ID, "selexcel")).select_by_visible_text('ReSubmissions')

        if is_rpo:
            driver.find_element(By.XPATH, "//*[@id='event-container']/form/fieldset/table/tbody/tr[2]/td[4]/a").click()
        else:
            driver.find_element(By.XPATH, "//*[@id='event-container']/form/fieldset/table/tbody/tr[2]/td[3]/a").click()
            
def rpo2(is_rpo=True):
    for j in range(len(dates)-1):
        for i in range(len(usernames1)):
            if j == 0:
                dtfrm = (dates[j+1] + timedelta(days=1)).strftime("%m/%d/%Y")
                dtto = dates[j].strftime("%m/%d/%Y")
            elif j > 0:
                dtfrm = dates[j+1].strftime("%m/%d/%Y")
                dtto = dates[j].strftime("%m/%d/%Y")
                
            login_and_filling(usernames1[i], passwords1[i], dtfrm, dtto, is_rpo)

def rpo1(is_rpo=True):
    for i in range(len(usernames2)):
        login_and_filling(usernames2[i], passwords2[i], datefrm1, dateto1, is_rpo)

def dhpo(is_rpo=False):
    for i in range(len(usernames3)):
        login_and_filling(usernames3[i], passwords3[i], datefrm1, dateto1, is_rpo)

if choice == 1:
    dhpo()
    cfx.show_info("Complete", "DHPO Completed")
elif choice == 2:
    rpo2()
    rpo1()
    cfx.show_info("Complete", "RPO Completed")
elif choice == 3:
    rpo1()
    cfx.show_info("Complete", "RPO 1 Completed")
elif choice == 4:
    rpo2()
    cfx.show_info("Complete", "RPO 2 Completed")
elif choice == 5:
    dhpo()
    rpo1()
    rpo2()
    cfx.show_info("Complete", "All Completed")
else:
    sys.exit()


