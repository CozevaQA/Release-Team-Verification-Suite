import re
import traceback
from datetime import date, datetime, time
import random

import pytz
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, \
    ElementClickInterceptedException
import time
import py_compile



def date_time():
    today = date.today()
    tz_In = pytz.timezone('Asia/Kolkata')
    datetime_In = datetime.now(tz_In)
    time = datetime_In.strftime("[%I-%M-%S %p]")
    now = str(today) + time
    return now



def ajax_preloader_wait1(driver):
    time.sleep(1)
    #WebDriverWait(driver, 300).until(
    #    EC.invisibility_of_element((By.XPATH, "//div/div[contains(@class,'ajax_preloader')]")))
    WebDriverWait(driver, 300).until(
        EC.invisibility_of_element((By.CLASS_NAME, "ajax_preloader")))
    WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.CLASS_NAME, "ajax_preloader hide")))

    time.sleep(1)


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

    WebDriverWait(driver, 300).until(
        EC.invisibility_of_element((By.CLASS_NAME, "drupal_message_text")))

    time.sleep(1)

    if check_exists_by_xpath(driver, "//*[@id='announcement_list']"):
        time.sleep(1)
        try:
            driver.find_element(By.XPATH, "//a[@OnClick='closeAnnouncements();']").click()
        except Exception as e:
            print("Announcement not present")


def ajax_preloader_wait2(driver):
    wait_time = 300
    time.sleep(1)
    loader_start_time = time.perf_counter()
    while time.perf_counter() - loader_start_time < wait_time:
        try:
            inner_content = driver.find_element_by_class_name("ajax_preloader").get_attribute("innerHTML")
            print(inner_content)
            if inner_content == "":
                time.sleep(1)
                break
        except NoSuchElementException as e:
            #traceback.print_exc()
            time.sleep(1)
            break
        except Exception as e:
            traceback.print_exc()
            time.sleep(1)
            continue

    return




def saml_toast_wait(driver):
    x=0




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
            try:
                toast_message = driver.find_elements_by_xpath(toast_message_xpath_new)[0].text
            except IndexError as e:
                toast_message = [""]
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

def action_click(element, driver):
    try:
        element.click()
    except (ElementNotInteractableException, ElementClickInterceptedException):
        driver.execute_script("arguments[0].click();", element)


def captureScreenshot(driver, page_title, screenshot_path):

    try:
        date = datetime.now().strftime('%H_%M_%S_%p')

        bad_chars = [';', ':', '|', ' ', '/', '\\']
        for i in bad_chars:
            final_title_text = page_title.replace(i, '_')

        driver.save_screenshot(screenshot_path + "/" + final_title_text + "_" + str(date) + ".png")
        #driver.save_screenshot(screenshot_path + "/" + page_title + "_" + str(date) + ".png")

    except Exception as e:
        print(e)

def URLAccessCheck(targetpath,driver):
    current_url = targetpath
    access_message = CheckAccessDenied(current_url)
    if access_message == 1:
        print("Access Denied found!")
        return True
    else:
        print("Access Check done!")
        error_message = CheckErrorMessage(driver)
        if error_message == 1:
            print("Error toast message is displayed")
            return True
            # logger.critical("ERROR TOAST MESSAGE IS DISPLAYED!")
        else:
            return False

def get_patient_id(href):
    cozeva_id = re.search('/patient_detail/(.*)?session', href)
    return (cozeva_id.group(1).replace("?", ""))
