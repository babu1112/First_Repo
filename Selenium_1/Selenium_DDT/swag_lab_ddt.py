from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from datetime import time
import time
from time import sleep
import XLUtils

ser = Service(executable_path=r"C:\Eclipse Workspace\chromedriver-win32\chromedriver.exe")
dr = webdriver.Chrome(service=ser)

dr.maximize_window()
dr.get("https://www.saucedemo.com/")
time.sleep(2)

path = r"D:\Selenium excel files\test_login_saucedemo.xlsx"

row = XLUtils.getRowCount(path, "Sheet1")
print("total no of fow in the sheet =", row)
col = XLUtils.getColumnCount(path, "Sheet1")
print("total no of column in the sheet =", col)

for r in range(2, row+1):
    username = XLUtils.readData(path, "Sheet1", r, 1)
    password = XLUtils.readData(path, "Sheet1", r, 2)

    dr.find_element(By.ID, 'user-name').send_keys(username)
    time.sleep(1)
    dr.find_element(By.ID, 'password').send_keys(password)
    time.sleep(1)
    dr.find_element(By.ID, 'login-button').click()
    print(dr.title)

    if dr.title == "Swag Labs":
        print("Login Successful")
        XLUtils.writeData(path, "Sheet1", r, 3, "Test Passed - Login Successful")

        dr.find_element(By.ID, 'react-burger-menu-btn').click()
        time.sleep(1)
        dr.find_element(By.XPATH, "//a[@id='logout_sidebar_link']").click()

    else:
        print("Login Un-Successful")
        XLUtils.writeData(path, "Sheet1", r, 3, "Test Failed - Loging Un-Successful")

dr.quit()


