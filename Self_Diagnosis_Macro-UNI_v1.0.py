"""
<Self_Diagnosis_Macro(UNI)> 22.3.3. (THU) 00:03
* Made by Yoonmen *
"""

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import chromedriver_autoinstaller
import time
import csv
from selenium.webdriver.common.alert import Alert
import pyautogui

# << Default Setting (1/5) >> --------------------

option = Options()
option.add_argument("start-maximized")


# << Read CSV File (2/5) >> --------------------

## Read CSV
CSV_file = open("./DB/user.csv", "r", encoding = "utf-8")
CSV_read = csv.reader(CSV_file)
CSV_data = []

## CSV(user.csv) -> List(CSV_data)
for row in CSV_read : 
    CSV_data.append(row)

## Ignore header
del CSV_data[0]
userNum = len(CSV_data)


# << chromedriver ON (3/5) >> --------------------

chromeVer = chromedriver_autoinstaller.get_chrome_version().split(".")[0]

try:
    driver = webdriver.Chrome(f"./{chromeVer}/chromedriver.exe", options=option)

except:
    chromedriver_autoinstaller.install(True)
    driver = webdriver.Chrome(f"./{chromeVer}/chromedriver.exe", options=option)

driver.implicitly_wait(1)


# << Start self-diagnosis (4/5) >> --------------------

for i in range(userNum) : 
    ## (자가진단 페이지 입장)
    driver.get("https://onestop.kumoh.ac.kr/html/etc/w_etc_survey_selfchk.html")
    time.sleep(10)

    ## (시작 ~ 사용자 정보 입력)
    pyautogui.write(CSV_data[i][1])
    driver.find_element_by_xpath("//*[@id=\"Form_비밀번호.비밀번호\"]").send_keys(CSV_data[i][2])
    driver.find_element_by_xpath("//*[@id=\"Form_로그인\"]").click()

    ## (자가진단 항목 체크)
    driver.find_element_by_xpath("//*[@id=\"Form_문항1.radio1.m1\"]").click()
    driver.find_element_by_xpath("//*[@id=\"Form_문항2.check2_1\"]").click()
    driver.find_element_by_xpath("//*[@id=\"Form_문항3.radio3.m1\"]").click()
    driver.find_element_by_xpath("//*[@id=\"Form_문항4.radio4.m1\"]").click()
    driver.find_element_by_xpath("//*[@id=\"Form_문항5.radio5.m1\"]").click()

    ## (제출)
    driver.find_element_by_xpath("//*[@id=\"Form8\"]").click()
    driver.switch_to_alert
    Alert(driver).accept()
    driver.switch_to_alert
    Alert(driver).accept()


# << End self-diagnosis (5/5) >> --------------------

driver.close()