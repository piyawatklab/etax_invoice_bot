from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from seleniumbase import Driver
import time
from logging import ERROR

import os
import glob

from libs.settings import USERNAME_LOGIN , PASSWORD_LOGIN , SHEET_LINK

PWD = os.getcwd()

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {"profile.default_content_settings.popups": 0,
            "download.default_directory": f'{PWD}/downloaded_files/',
            "directory_upgrade": True}
options.add_experimental_option("prefs",prefs)

data = []

def run():

    # 1. ลบ invoice - main.csv เดิมออก
    delete_invoice()

    # 2. Download invoice - main.csv ตัวใหม่มาเก็บไว้
    driver = Driver(uc=True)
    url = SHEET_LINK
    driver.get(url)
    time.sleep(3)

    # 3. อ่านเลขที่ Invoice จาก invoice - main.csv
    invoice_list = search_invoice()

    # 4. เริ่มต้น Download Invoice
    driver.set_window_size(1920, 1080)

    # 4.1 Open ETAX Webapp
    url = 'https://etax.invchain.com/invoicechain/#/login'
    driver.get(url)
    time.sleep(3)

    # 4.2 Login to ETAX
    driver.find_element(By.XPATH, "//input[@name='username']").send_keys(USERNAME_LOGIN)
    driver.find_element(By.XPATH, "//input[@name='password']").send_keys(PASSWORD_LOGIN)
    time.sleep(1)
    driver.find_element(By.XPATH, "//button[@name='btn_login']").click()
    time.sleep(3)
    driver.find_element(By.XPATH, "//button[@class='swal2-confirm btn btn-info btn-fill']").click()
    time.sleep(3)

    driver.find_element(By.XPATH, "//a[@name='showSearchForm']").click()
    time.sleep(1)

    for line in invoice_list:

        try:
            # 4.3 Search invoice no. from invoice list
            search_doc = driver.find_element(By.XPATH, "//input[@name='documentNo']")
            driver.execute_script("arguments[0].scrollIntoView(true);", search_doc)

            search_doc.clear()
            search_doc.send_keys(line)
            
            driver.find_element(By.XPATH, "//button[@name='doSearch']").click()
            time.sleep(3)

            invoice_search = driver.find_element(By.TAG_NAME, "tbody")
            invoice_search_line = invoice_search.find_elements(By.TAG_NAME, "tr")

            # 4.4 Download invoice
            if len(invoice_search_line) != 0 :
                driver.execute_script("arguments[0].scrollIntoView(true);", invoice_search_line[0])
                invoice_search_line[0].find_element(By.XPATH, "//a[@name='downloadPDF0']").click()
            
            time.sleep(3)

        except Exception as e:
            time.sleep(1)

    driver.quit()

def download_invoice():

    driver = Driver(uc=True)

    url = SHEET_LINK
    driver.get(url)

    time.sleep(3)
    driver.quit()

def search_invoice():

    with open('downloaded_files/invoice - main.csv', 'r') as file:
        csv_data = file.readlines()

    data_array = []
    for line in csv_data:
        line = line.strip()
        if line == 'invoice':
            continue
        data_array.append(line)

    return data_array
    
def delete_invoice():
    files = glob.glob(f'downloaded_files/*csv')
    for f in files:
        os.remove(f)

if __name__ == "__main__":

    run()