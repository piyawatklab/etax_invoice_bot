from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from seleniumbase import Driver
import time
from logging import ERROR

import os
import glob
import shutil
import sys

import pandas as pd

from libs.settings import USERNAME_LOGIN , PASSWORD_LOGIN , SHEET_LINK , DESTINATION_PATH

PWD = os.getcwd()

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {"profile.default_content_settings.popups": 0,
            "download.default_directory": f'{PWD}/downloaded_files/',
            "directory_upgrade": True}
options.add_experimental_option("prefs",prefs)

data = []

def run(name):

    # 1. เปิด Browser ขึ้นมา
    # driver = Driver(uc=True)
    driver = Driver(uc=True,headless=True)

    # 2. ตรวจสอบการรัน
    if name == 'online':

        # 2.1 ถ้ารันแบบ online ลบ invoice.xlsx เดิมออก
        clean_order()

        # 2.2 โหลดไฟล์ invoice.xlsx ใหม่มาเก็บไว้
        url = SHEET_LINK
        driver.get(url)
        time.sleep(3)

    # 3. อ่านเลขที่ Invoice จาก invoice.xlsx
    invoice_list = search_invoice()
    
    # 4. เริ่มต้น Download Invoice
    if len(invoice_list) > 0 :
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
                print(line,'OK')

            except Exception as e:
                time.sleep(1)
                print(line,'Error')

    driver.quit()
    print('Done')
    time.sleep(3)
    move_file()

def search_invoice():

    df = pd.read_excel('downloaded_files/invoice.xlsx')

    data_array = df.values.tolist()
    data_array = [str(item[0]) for item in data_array]

    print(data_array)

    return data_array

def clean_order():

    # files = glob.glob(f'downloaded_files/*pdf')
    # for f in files:
    #     os.remove(f)

    source_directory = 'downloaded_files'
    destination_directory = 'downloaded_files/backup_files'

    if not os.path.exists(destination_directory):
        os.makedirs(destination_directory)

    files = os.listdir(source_directory)
    for file_name in files:
        if file_name.endswith('.pdf') :
            source_path = os.path.join(source_directory, file_name)
            destination_path = os.path.join(destination_directory, file_name)
            shutil.move(source_path, destination_path)

def move_file():
    try : 
        files = glob.glob(f'downloaded_files/*pdf')

        for source_file in files:
            file_name = os.path.basename(source_file)
            print(file_name)
            new_file_name = file_name.split('_')[0] + ".pdf"
            new_path = DESTINATION_PATH
            # shutil.move(source_file, f'{new_path}\\{file_name}')
            shutil.move(source_file, f'{new_path}\\{new_file_name}')
            
    except Exception as e:
        time.sleep(1)
        print(e, 'move_file Error')

if __name__ == "__main__":

    name = 'default'
    if len(sys.argv) > 1:
        for arg in sys.argv[1:]:
            key, value = arg.split("=")
            if key == "name":
                name = value
    
    print("Name :", name)
    run(name)
    
    