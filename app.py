from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from win10toast import ToastNotifier
from datetime import datetime
import time

# ############################# Constants Config #######################################
EMAIL_ID = "sam@wavelabs.ai"
PASSWORD = "nuvafiw*97"


# #####################################################################################

# ################################## Main program starts ####################################
def showToast(title, content, duration):
    toast.show_toast(
        title,
        content,
        duration=duration,
        icon_path="icon.ico",
        threaded=True,
    )


chrome_options = Options()
chrome_options.add_argument("--disable-user-media-security=true")

driver = webdriver.Chrome(executable_path='chromedriver.exe', chrome_options=chrome_options)
toast = ToastNotifier()

driver.get("https://accounts.zoho.in/signin?servicename=zohopeople&signupurl=https://www.zoho.in/people/signup.html")

# username field
try:
    check_page_load = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "login_id")),
        EC.element_to_be_clickable((By.ID, "nextbtn"))
    )
finally:
    email = driver.find_element_by_id('login_id')
    email.send_keys(EMAIL_ID)
    nextBtn = driver.find_element_by_id('nextbtn')
    nextBtn.click()

# password field
try:
    check_page_load = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.ID, "password"))
    )
finally:
    password = driver.find_element_by_id('password')
    password.send_keys(PASSWORD)
    nextBtn = driver.find_element_by_id('nextbtn')
    nextBtn.click()

try:
    check_page_load = WebDriverWait(driver, 100).until(
        EC.presence_of_element_located((By.ID, "zp_maintab_home"))
    )

finally:
    homeBtn = driver.find_element_by_id('zp_maintab_home')
    homeBtn.click()
    try:
        check_page_load = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "ZPD_Top_Att_Stat"))
        )
    finally:
        # ZPD_Top_Att_Stat
        checkInBtnColorProp = driver.find_element_by_id("ZPD_Top_Att_Stat").get_attribute("class")
        now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        if checkInBtnColorProp == "in CP":
            # CHECK-IN
            driver.execute_script("Attendance.Dashboard.CurrStatus.updateCheckOut(true)")
            print("checked-in")
            showToast("Attendance Checked-in", "Attendance for wavelabs was checkedin at " + now, 20)

        else:
            # CHECK-OUT
            driver.execute_script("Attendance.Dashboard.CurrStatus.updateCheckOut(false)")
            print("checked-out")
            showToast("Attendance Checked-out", "Attendance for wavelabs was checkedout at " + now, 20)

time.sleep(10)
driver.close()
driver.quit()
