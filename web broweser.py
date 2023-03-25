from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from openpyxl import Workbook
import datetime

date = datetime.datetime.now().strftime("%d-%m-%Y")

wb = Workbook()
ws = wb.active

driver_path = r"E:\Python project\Google maps contact details\Driver\chromedriver_win32\chromedriver.exe"

url = r"https://www.google.com/maps"

search_data = "Used appliances Dealers in Karnataka "

driver = webdriver.Chrome(executable_path=driver_path)
driver.get(url)
time.sleep(5)

driver.find_element(By.ID,"searchboxinput").send_keys(search_data)
time.sleep(3)
driver.find_element(By.ID,"searchbox-searchbutton").click()
time.sleep(3)
l = 2
for r in range(1,40):
    print(r)

    Scroll = driver.find_element(By.CLASS_NAME, "hfpxzc")
    Scroll.send_keys(Keys.PAGE_DOWN)
    time.sleep(2)

    # if driver.find_element(By.CLASS_NAME,"m6QErb").text ""


for link in driver.find_elements(By.CLASS_NAME, "hfpxzc"):
    print(link.get_attribute("href"))
    ws.cell(row=l, column=1).value = link.get_attribute("href")
    l = l+1

wb.save(r"E:\Python project\Google maps contact details\Excel\ " + search_data + " " + date + ".xlsx")

