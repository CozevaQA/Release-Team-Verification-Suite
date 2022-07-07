#todo navigate to workbook
#click on download icon
#click on download workbook data
#extract the properties of file

#folder structure customerid/worksheet/year/lob/file
#report - HTML Report Worksheet comparison
#format - Worksheet name as table name ; parameters as reports
import os
import shutil

import selenium
import configparser
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, \
    ElementClickInterceptedException
from selenium.webdriver import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
config = configparser.RawConfigParser()
config.read("locator-config.properties")

def wait_to_load(driver):
    loader=config.get("PharmacyCost-Prod","loader_element")
    WebDriverWait(driver,100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader)))

def action_click(driver,element):
    try:
        element.click()
    except (ElementNotInteractableException, ElementClickInterceptedException):
        driver.execute_script("arguments[0].click();", element)


def copy_paste_file(download_directory,report_directory):
    files = os.listdir(download_directory)
    print(files)
    shutil.copytree(download_directory, report_directory)

def click_on_download(driver,worksheet):
    download_icon_xpath=config.get("runner","download_icon_xpath")
    download_icon=driver.find_element_by_xpath(download_icon_xpath)
    action_click(driver,download_icon)
    # WebDriverWait(driver, 30).until(
    #     EC.visibility_of_element_located((By.CLASS_NAME, config.get("runner","dropdown_table") )))
    print("Clicked on Download ")
    if "Cohort" in worksheet:
        download_all_data=config.get("runner","download_all_data_xpath_cohort")
    else:
        download_all_data=config.get("runner","download_all_data_xpath_other")

    download_all_data_element=driver.find_element_by_xpath(download_all_data)
    action_click(driver,download_all_data_element)


def download_workbook_data(driver,download_directory,report_directory,customerid,worksheet,year,LOB):
    click_on_download(driver,worksheet)


