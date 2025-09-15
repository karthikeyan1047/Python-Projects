import pyautogui as pag
import pygetwindow as gw
import time, clipboard, sys, os
import _functions as cfx

def wait_for_window(title_substring, filename, timeout=180):
    start_time = time.time()
    while time.time() - start_time < timeout:
        windows = gw.getWindowsWithTitle(title_substring)
        if windows:
            clipboard.copy(filename)
            time.sleep(1)
            pag.hotkey('ctrl', 'v')
            time.sleep(1)
            pag.press('enter')
            time.sleep(1)
            pag.press('y')
            return True
        time.sleep(1)
    return False

def download_wait(path, timeout=120, extra_wait=3):
    start_time = time.time()
    while not os.path.exists(path):
        if time.time() - start_time > timeout:
            return False
        time.sleep(1)
    time.sleep(extra_wait)
    return True


choice = int(cfx.inputbox("Select", "1. DHPO\n2. RPO"))
choice_map = {
    1 : "DHPO",
    2 : "RPO"
}
ch = choice_map.get(choice, "Invalid")

if ch == "DHPO":
    x,y = 1485, 348       
    # x,y = 1485, 298
    ext = '.xlsx'
    names = ["XML_DRS_SMCL_DHA", "XML_DRS_JOC", "XML_DRS_HOC", "XML_DRS_med", "XML_DRS_simc", 
             "XML_DRS_SMCL_MOH", "XML_DRS_GOC", "DRS_DHPO_IPOP", "DRS_DHPO_PH"]

elif ch == "RPO":
    x,y = 1464, 343       
    # x,y = 1464, 293
    ext = '.csv'
    names = ["DRS_RPO_L153", "DRS_RPO_L5223", "DRS_RPO_L5741", "DRS_RPO_L7026", "DRS_RPO_L5887", "DRS_RPO_L5136", 
            "DRS_RPO_L206 1", "DRS_RPO_IPOP 1", "DRS_RPO_PH 1", "DRS_RPO_L206 2", "DRS_RPO_IPOP 2","DRS_RPO_PH 2"]
    
else:
    sys.exit(0)

download_dir = r"C:\Users\karthikeyans\Downloads"
for name in names:
    op_filepath = os.path.join(download_dir, f"{name}{ext}")
    time.sleep(1)
    pag.click(x,y)
    wait_for_window("Save As", name, 120)
    if download_wait(path=op_filepath):
        pag.hotkey('ctrl', 'w')

