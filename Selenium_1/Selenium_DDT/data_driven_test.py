# from telnetlib import EC
# from selenium.webdriver.support.wait import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC

import XLUtils

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from datetime import time
from time import sleep

ser = Service(executable_path=r"C:\Eclipse Workspace\chromedriver-win32\chromedriver.exe")
dr = webdriver.Chrome(service=ser)
dr.maximize_window()
# Set implicit wait time (e.g., 10 seconds)
# dr.implicitly_wait(10)

dr.get("https://demo.guru99.com/test/newtours/")
dr.implicitly_wait(10)

path = "D:\\Selenium excel files\\test_login.xlsx"

row = XLUtils.getRowCount(path, 'Sheet1')
print("no of rows = ", row)
col = XLUtils.getColumnCount(path, 'Sheet1')
print("no of columns = ", col)

for r in range(2, row + 1):
    username = XLUtils.readData(path, "Sheet1", r, 1)
    password = XLUtils.readData(path, "Sheet1", r, 2)

    dr.find_element(By.NAME, 'userName').send_keys(username)
    dr.implicitly_wait(2)
    dr.find_element(By.NAME, 'password').send_keys(password)
    dr.implicitly_wait(2)
    dr.find_element(By.XPATH, "//input[@name='submit']").click()
    dr.implicitly_wait(2)
    # Wait until the username field is present
    '''WebDriverWait(dr, 10).until(EC.presence_of_element_located((By.NAME, 'userName'))).send_keys(username)
    sleep(2)
    WebDriverWait(dr, 10).until(EC.presence_of_element_located((By.NAME, 'password'))).send_keys(password)
    sleep(2)
    dr.find_element(By.XPATH, "//input[@name='submit']").click()
    sleep(2)'''

    if dr.title == "Login: Mercury Tours":
        print("Login is Successful by USER")
        XLUtils.writeData(path, "Sheet1", r, 3, "Test Passed")
        print("data written into excel file --> if")
        dr.find_element(By.LINK_TEXT, 'Home').click()
    else:
        print("Login is un-successful by USER")
        XLUtils.writeData(path, "Sheet1", r, 3, "Test Failed")
        print("data written into excel file --> else")

    # dr.find_element(By.LINK_TEXT, 'Home').click()
    dr.sleep(2)

dr.quit()
