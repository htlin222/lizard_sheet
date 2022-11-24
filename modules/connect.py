import pickle
import ctypes
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup

# setup
options = Options()
options.add_argument("--disable-notifications")
driver = webdriver.chrome(chrome_options=options)

# LOGIN_NAME = 'DOC8633F'
# PASSWORD = '90-op[90-'
# Login
# https://weirenxue.github.io/2021/10/29/selenium_tkinter_login/
driver.get('https://eip.vghtpe.gov.tw/login.php')
time.sleep(15)
response = ctypes.windll.user32.MessageBoxW(0, "你登入成功了嗎?", "注意，是的話再按確定", 1)

if response ==1:
    driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findEmr&histno=46784901") #先連到我的病歷看一眼建立session
    driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findPatient")
    cookies = driver.get_cookies()
    with open('../local/cookies.json','wb') as f:
        pickle.dump(cookies,f)
    print ('done')
else:
    driver.quit()

# 之後的登入
def get_html_from_url_with_cookie(url):
    '''
    as title, need cookie files
    '''
    options = Options()
    options.headless = True
    options.add_argument("--disable-notifications")
    driver = webdriver.chrome(chrome_options=options)
    driver.get(url)
    #刪掉cookies
    driver.delete_all_cookies()

    with open('D:/test_cookies/db_cookie_1','rb') as cookie_json:
        loaded_cookies = pickle.load(cookie_json)
    for cookie in loaded_cookies:
        driver.add_cookie(cookie)
        print(cookie)
    #帶我們保存的cookie訪問豆瓣
    driver.get(url)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    print('done')
    return soup
