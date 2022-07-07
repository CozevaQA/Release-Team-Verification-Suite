from datetime import date, datetime, time
import random

import pytz
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import time
import py_compile


def date_time():
    today = date.today()
    tz_In = pytz.timezone('Asia/Kolkata')
    datetime_In = datetime.now(tz_In)
    time = datetime_In.strftime("[%I-%M-%S %p]")
    now = str(today) + time
    return now


def ajax_preloader_wait(driver):
    time.sleep(1)
    #WebDriverWait(driver, 300).until(
    #    EC.invisibility_of_element((By.XPATH, "//div/div[contains(@class,'ajax_preloader')]")))
    WebDriverWait(driver, 300).until(
        EC.invisibility_of_element((By.CLASS_NAME, "ajax_preloader")))
    #time.sleep(1)
    if len(driver.find_elements_by_class_name("ajax_preloader")) != 0:
        WebDriverWait(driver, 300).until(
            EC.invisibility_of_element((By.CLASS_NAME, "ajax_preloader")))
    time.sleep(1)




def CheckAccessDenied(string):
    sub_str = "/access_denied"
    if string.find(sub_str) == -1:
        # print("Access check done")
        return 0
    else:
        print("ACCESS DENIED has been found!!")
        return 1


def CheckErrorMessage(driver):
    err_msg = 0
    sub_str = "error"
    toast_messages = driver.find_elements_by_xpath("//div[@class='drupal_message_text']")
    if (len(toast_messages)) != 0:
        i = 1
        while i <= len(toast_messages):
            toast_message_xpath_new = "(//div[@class='drupal_message_text'])" + str([i])
            toast_message = driver.find_element_by_xpath(toast_message_xpath_new).text
            if toast_message.count(sub_str) > 0:
                err_msg = 1
                break
            i += 1
        if err_msg == 1:
            return 1
        else:
            return 0
    else:
        return 0


def RandomNumberGenerator(maximum_range,number):
    a = []
    a = random.sample(range(1, maximum_range), number)
    #print(a)
    return a

def check_exists_by_xpath(driver, xpath):
    try:
        driver.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

def check_exists_by_css(driver, css):
    try:
        driver.find_element_by_css_selector(css)
    except NoSuchElementException:
        return False
    return True

def check_exists_by_class(driver, classname):
    try:
        driver.find_element_by_class_name(classname)
    except NoSuchElementException:
        return False
    return True

def check_exists_by_id(driver, id):
    try:
        driver.find_element_by_id(id)
    except NoSuchElementException:
        return False
    return True