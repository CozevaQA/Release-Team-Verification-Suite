import setups
import logging
import ExcelProcessor as db
import context_functions as cf
import support_functions as sf
import setups as st
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import variablestorage as locator
import openpyxl
if __name__ == '__main__':
    print("Hello World")
    driver = setups.driver_setup()
    setups.login_to_cozeva()
    sf.ajax_preloader_wait(driver)

    excel_path = ""



    def performGlobalSearch(username, keywords):




