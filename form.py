import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook,load_workbook
import datetime
date = datetime.datetime.now().strftime("%d-%m-%Y")
driver_path = r"E:\Python project\Google maps contact details\Driver\chromedriver_win32\chromedriver.exe"
web_url = r"https://www.google.com/maps"
extension = ".xlsx"
location_save = r"E:\Python project\Google maps contact details\Excel"
driver = webdriver.Chrome(executable_path=driver_path)

wb = Workbook()
ws = wb.active


class GoogleMaps:
    def browser_run(self):
        driver.get(web_url)

    def input_fields(self,  **kwargs):
        driver.find_element(By.ID, "searchboxinput").send_keys(kwargs.get("find"))
        time.sleep(3)

    def input_fields_post(self):
        driver.find_element(By.ID, "searchbox-searchbutton").click()
        time.sleep(3)

    def scroll(self):
        for r in range(1, 3):
            print(r)
            driver.find_element(By.CLASS_NAME, "hfpxzc").send_keys(Keys.PAGE_DOWN)
            time.sleep(2)

    def get_link(self, r=2):
        for link in driver.find_elements(By.CLASS_NAME, "hfpxzc"):
            print(link.get_attribute("href"))
            ws.cell(row=r, column=1).value = link.get_attribute("href")
            r = r + 1
        self.save_excel(title="\data ")

    def save_excel(self,**kwargs):
        wb.save(location_save + kwargs.get("title") + date + extension)


    def search_text(self, **kwargs):
        print(kwargs.get("searchText"), "in", kwargs.get("Location"))
        findtext = kwargs.get("searchText"), " in ", kwargs.get("Location")
        self.browser_run()
        self.input_fields(find=findtext)
        self.input_fields_post()
        self.scroll()
        self.get_link()
        time.sleep(3)
        self.get_data()

    def get_phone(self):
        for ph in driver.find_elements(By.TAG_NAME, "button"):
            if len(ph.text) == 12:
                return ph.text

    def element(self,**kwargs):
        try:
            name = driver.find_element(kwargs.get("by"), kwargs.get("value")).text
            self.element_print(heading=kwargs.get("heading"), value=name)
            return name
        except:
            return None

    def element_print(self,**kwargs):
        print(kwargs.get("heading"), ":", kwargs.get("value"))

    def maps_data_range(self):
        for r in range(2, 15): #130
            print(r)

            driver.get(ws.cell(row=r, column=1).value)
            name = self.element(heading="Name", by=By.TAG_NAME, value="h1")
            address = self.element(heading="Address", by=By.CLASS_NAME, value="Io6YTe")
            website = self.element(heading="Website", by=By.CLASS_NAME, value="ITvuef")
            phone = self.get_phone()

            print("Phone :", phone)

            ws.cell(row=r, column=2).value = name
            ws.cell(row=r, column=3).value = address
            ws.cell(row=r, column=4).value = website
            ws.cell(row=r, column=5).value = phone

    def save_excel_details(self):
        wb.save(location_save + "\data_details " + date + extension)

    def get_data(self):
        print(location_save + "\data " + date + extension)
        lb = load_workbook(location_save + "\data " + date + extension)
        ls = lb.active
        self.browser_run()
        self.maps_data_range()
        self.save_excel(title="\data_details ")




