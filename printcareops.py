import time
import traceback
import pyautogui

from openpyxl.styles import Font, PatternFill
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

import setups
import logging
import ExcelProcessor as db
import context_functions as cf
import support_functions as sf
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import variablestorage as locator
from openpyxl import Workbook, load_workbook
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, \
    ElementClickInterceptedException, TimeoutException

ENV = 'CERT'
ID = "1000"

def cancelPrintPreview():
    # get the current time and add 180 seconds to wait for the print preview cancel button
    endTime = time.time() + 180
    # switch to print preview window
    driver.switch_to.window(driver.window_handles[-1])
    while True:
        try:
            # get the cancel button
            cancelButton = driver.execute_script(
                "return document.querySelector('print-preview-app').shadowRoot.querySelector('#sidebar').shadowRoot.querySelector('print-preview-header#header').shadowRoot.querySelector('paper-button.cancel-button')")
            if cancelButton:
                # click on cancel
                cancelButton.click()
                # switch back to main window
                driver.switch_to.window(driver.window_handles[0])
                return True
        except:
            pass
        time.sleep(1)
        if time.time() > endTime:
            driver.switch_to.window(driver.window_handles[0])
            break

if __name__ == '__main__':
    print("Hello World")
    driver = setups.driver_setup()
    if ENV == 'CERT':
        setups.login_to_cozeva_cert(ID)
    elif ENV == 'STAGE':
        setups.login_to_cozeva_stage()
    elif ENV == "PROD":
        setups.login_to_cozeva(ID)
    else:
        print("ENV INVALID")
        exit(3)

    sf.ajax_preloader_wait(driver)

    driver.get("https://cert.cozeva.com/patient_detail/168ZD6A?session=YXBwX2lkPXJlZ2lzdHJpZXMmcGF5ZXJJZD0xMDAwJmN1c3RJZD0xMDAwJm9yZ0lkPTEwMDA%3D&cozeva_id=168ZD6A&patient_id=8736791&tab_type=CareOps&first_load=1")
    sf.ajax_preloader_wait(driver)

    driver.find_element(By.CLASS_NAME, "patient_print_options").click()
    time.sleep(2)

    dropdown_elements = driver.find_element(By.ID, "patient_print_dropdown").find_elements(By.TAG_NAME, "li")
    print_options = ["Print Careops Summary", "Print HCC Confirm/Disconfirm","Print Careops","Print Med Adherence","Print HCC"]
    for element in dropdown_elements:
        print(element.text)
    print_window_open = 0
    for element in dropdown_elements:
        if element.text in print_options:
            try:
                print("Clicking on "+element.text)
                #element.click()
                #driver.execute_script("arguments[0].click();", element)
                webdriver.ActionChains(driver).move_to_element(element).click(element).perform()
                #print(driver.window_handles)
                print("Clicked on "+element.text)
                time.sleep(1)
                print("Waiting done")
                #take a screenshot here using robot

                print(driver.window_handles)
                driver.switch_to.window(driver.window_handles[1])
                print("Switching window")
                print_window_open = 1
                print("Capturing Screenshot")
                sf.captureScreenshot(driver, element.text, locator.parent_dir)
                print("Screenshot captured")
                driver.close()
                driver.switch_to.window(driver.window_handles[0])


            except Exception as e:
                traceback.print_exc()
                if print_window_open == 1:
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])


    time.sleep(2)

