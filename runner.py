from pathlib import Path

import xlwt as xlwt

import setups
import variablestorage as vs
from selenium import webdriver
# import xlrd
import openpyxl as xlrd
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, WebDriverException, \
    ElementNotInteractableException, ElementClickInterceptedException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
# from QualityOverview import QualityOverview
# from TotalCost import TotalCost
# from EDCost import EDCost
# from InpatientCost import InpatientCost
import sqlite3
from sqlite3 import Error
import sys
import base64
import os
from os import path
import shutil
# from SendReport import SendReport
# from MedicareRiskOverview import MedicareRiskOverview
# from SuspectAnalytics import SuspectAnalytics
# from MedicareCodingDiscontinuation import MedicareCodingDiscontinuation
import configparser
from plyer import notification
# from PharmacyCost import PharmacyCost
# #from InpatientCost import InpatientCost
# from MedicareRAF import MedicareRAF
# from MedicareRiskByHierarchy import MedicareRiskByHierarchy
# from CommercialCodingDiscontinuation import CommercialCodingDiscontinuation
# from CommercialRiskOverview import CommercialRiskOverview
# from InpatientCostByDX import InpatientCostByDX
# from InpatientCostByHCC import InpatientCostByHCC
config = configparser.RawConfigParser()
config_path = Path("locator-config.properties")
config.read(config_path)
# from ProfessionalCost import ProfessionalCost
# from TotalCostTrends import TotalCostTrends
# from EDCostTrends import EDCostTrends
# from CohortAnalyzer import CohortAnalyzer
# from CohortAnalyzerSummary import CohortAnalyzerSummary
import pickle
#import xxx

import logging
import os
import shutil
from datetime import date
from datetime import date, datetime
import pytz
import configparser


# from UsageMonthlyActivity import UsageMonthlyActivity

# def get_logger(
#         LOG_FORMAT     = '%(asctime)s %(name)-12s %(levelname)-8s %(message)s',
#         LOG_NAME       = 'Runner',
#         LOG_FILE_INFO  = os.getcwd()+"\\"+str(master_directory)+"\\"+str(path)+"\\"+'file.log',
#         LOG_FILE_ERROR = os.getcwd()+"\\"+str(path)+"\\"+'Error.log'):
#
#     log           = logging.getLogger(LOG_NAME)
#     log_formatter = logging.Formatter(LOG_FORMAT)
#
#     # comment this to suppress console output
#     stream_handler = logging.StreamHandler()
#     stream_handler.setFormatter(log_formatter)
#     log.addHandler(stream_handler)
#
#     file_handler_info = logging.FileHandler(LOG_FILE_INFO, mode='w')
#     file_handler_info.setFormatter(log_formatter)
#     file_handler_info.setLevel(logging.INFO)
#     log.addHandler(file_handler_info)
#
#     file_handler_error = logging.FileHandler(LOG_FILE_ERROR, mode='w')
#     file_handler_error.setFormatter(log_formatter)
#     file_handler_error.setLevel(logging.ERROR)
#     log.addHandler(file_handler_error)
#
#     log.setLevel(logging.INFO)
#
#     return log
#


def setup(val, downloaddefault):
    # #val.lower() == "firefox":
    # driver = webdriver.Firefox(executable_path=r"C:\\Users\\ssrivastava\\PycharmProjects\\Python_Practise\\driver2\\geckodriver.exe")
    # driver.implicitly_wait(10)
    # title = driver.get("https://stage.cozeva.com/user/login")
    # driver.maximize_window()
    if val.lower() == "chrome":
        options = webdriver.ChromeOptions()
        options.add_argument("--disable-notifications")
        options.add_argument("--start-maximized")
        options.add_argument(vs.chrome_profile_path)
        preferences = {
            "download.default_directory": downloaddefault}
        options.add_experimental_option("prefs", preferences)
        # Path to your chrome profile
        # options.add_argument("--headless")
        # options.add_argument('--disable-gpu')
        # options.add_argument("--window-size=1920,1080")
        # options.add_argument("--start-maximized")
        print("helloo")
        print(os.getcwd())
        global driver
        driver = webdriver.Chrome(executable_path=source_directory + "\\" + vs.chrome_driver_path, options=options)
        # preferences = {
        #     "download.default_directory":downloaddefault}
        # options.add_experimental_option("prefs", preferences)
        # # change for user
        # global driver
        # driver = webdriver.Chrome(executable_path=locator.chrome_driver_path, options=options)
        driver.get("https://www.cozeva.com/user/logout")
        title = driver.get("https://www.cozeva.com/user/login")
        driver.maximize_window()
    return driver

def remove_element_if_present(element_xpath):
    logger.info("0 debug")
    element=driver.find_element_by_xpath(element_xpath)
    logger.info(element.is_displayed())
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.XPATH,element_xpath)))
    try:
        driver.execute_script("""var l = document.getElementById("cozeva_support_chat_dasboard");
    l.parentNode.removeChild(l);""")
    except Exception as e:
        print(e)
        logger.info(e)


def makedir(foldername):
    path1 = str(foldername)
    if not os.path.exists(path1):
        try:
            os.mkdir(path1)
            return path1
        except OSError as error:
            print(error)
            return False
    else:
        try:
            shutil.rmtree(path1)
            os.mkdir(path1)
            return path1
        except OSError as error:
            print(error)
            return False


def date_time():
    today = date.today()
    # print("Today's date:", today)
    tz_In = pytz.timezone('Asia/Kolkata')
    datetime_In = datetime.now(tz_In)
    # print("IN time:", datetime_In.strftime("%I;%M;%S %p"))
    time = datetime_In.strftime("[%I-%M-%S %p]")
    now = str(today) + time
    print(now)
    # logger.info("Date and Time captured!")
    return (now)


def wait_till_value(path, delay, value):
    try:
        WebDriverWait(driver, delay).until(EC.text_to_be_present_in_element((By.XPATH, path), value))
        print("Page ready")
    except TimeoutException:
        print("Loading is taking too much time")


def check_exists_byclass(driver, classname):
    try:
        driver.find_element_by_class_name(classname)
    except NoSuchElementException:
        return False
    return True


def create_connection(db_file):  # creating connection
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except Error as e:
        print(e)

    return None


def truncate_table(table_name):
    # change for user
    db_path = config.get("runner", "dbpath")
    print(db_path)
    folder_path = ''
    try:
        conn = create_connection(db_path)
        cur = conn.cursor()
        query = "DELETE FROM " + str(table_name)
        print("deleted")
        cur.execute(query)
        conn.commit()
    except sqlite3.IntegrityError as e:
        print(e)


def make_directory(customer):
    customer_id = customer
    path1 = str(customer_id)
    if not os.path.exists(path1):
        try:
            os.mkdir(path1)
            return path1
        except OSError as error:
            print(error)
            return False
    else:
        try:
            shutil.rmtree(path1)
            os.mkdir(path1)
            return path1
        except OSError as error:
            print(error)
            return False


def login(driver, loc):
    filepath = source_directory + "\\assets\\loginInfo.txt"
    file = open(filepath, "r+")
    details = file.readlines()
    driver.find_element_by_id("edit-name").send_keys(details[0].strip())
    driver.find_element_by_id("edit-pass").send_keys(details[1].strip())
    file.seek(0)
    file.close()
    driver.find_element_by_id("edit-submit").click()
    # reason for login
    WebDriverWait(driver, 120).until(
        EC.presence_of_element_located((By.XPATH, "//textarea[@id=\"reason_textbox\"]")))
    actions = ActionChains(driver)
    reason = driver.find_element_by_xpath("//textarea[@id=\"reason_textbox\"]")
    actions.click(reason)
    actions.send_keys_to_element(reason, "https://redmine2.cozeva.com/issues/16428 ")
    actions.perform()
    driver.find_element_by_id("edit-submit").click()


# def action_click(element):
#     webdriver.ActionChains(driver).move_to_element(element).click(element).perform()

def action_click(element):
    try:
        element.click()
    except (ElementNotInteractableException, ElementClickInterceptedException):
        driver.execute_script("arguments[0].click();", element)


def nav_back():
    loader_element = 'sm_download_cssload_loader_wrap'
    WebDriverWait(driver, 90).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
    back_xpath = "//i[text()=\"arrow_back\"]"
    back = driver.find_element_by_xpath(back_xpath)
    action_click(back)


# def verify_totalcost(year, customer_id):
#     str2 = "//td[@class=\"sm_tab_link\" and text()=\"%s\"]" % "Total Cost"
#     wb = driver.find_element_by_xpath(str2)
#     action_click(wb)
#     f = TotalCost(driver)
#     f.iterate_filter(year, customer_id)
#     loader_element = 'sm_download_cssload_loader_wrap'
#     WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
#     nav_back()


def verify_Usage(year, customer_id):
    try:
        str2 = "//td[@class=\"sm_tab_link\" and text()=\"%s\"]" % "Cozeva Usage Monthly Activity"
        wb = driver.find_element_by_xpath(str2)
        print(wb.text)
        logger.info("Cozeva Usage Monthly Activity")
        action_click(wb)
        year = [year]
        f = UsageMonthlyActivity(driver)
        try:
            f.iterate_filter(year, customer_id)
            loader_element = 'sm_download_cssload_loader_wrap'
            WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
            nav_back()
        except TimeoutException as e2:
            close_button_when_loading = driver.find_element_by_xpath(
                config.get("runner", "close_button_when_loading_xpath"))
            print("Time out exception for ", customer_id, " Cozeva Usage Monthly Activity")
            logger.error(
                str(e2) + str(customer_id) + " Timeout Exception occurred in" + "Cozeva Usage Monthly Activity" + "\n")
            action_click(close_button_when_loading)
            nav_back()
            pass
        except (
                WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
                StaleElementReferenceException) as e:
            print(e)
            print("Exception occurred in Utilization  " + "Cozeva Usage Monthly Activity")
            logger.error(str(e) + str(customer_id) + "Cozeva Usage Monthly Activity" + "\n")
            nav_back()
            pass
    except NoSuchElementException:
        pass


# def verify_pharmacycost(year, customer_id):
#     str2 = "//td[@class=\"sm_tab_link\" and text()=\"%s\"]" % "Pharmacy Cost"
#     wb = driver.find_element_by_xpath(str2)
#     action_click(wb)
#     f = PharmacyCost(driver)
#     f.iterate_filter(year, customer_id)
#     loader_element = 'sm_download_cssload_loader_wrap'
#     WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
#     nav_back()
# def verify_edcost(year, customer_id):
#     str2 = "//td[@class=\"sm_tab_link\" and text()=\"%s\"]" % "ED Cost"
#     wb = driver.find_element_by_xpath(str2)
#     action_click(wb)
#     f = EDCost(driver)
#     f.iterate_filter(year, customer_id)
#     loader_element = 'sm_download_cssload_loader_wrap'
#     WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
#     nav_back()
#
# def verify_inpatientcost(year, customer_id):
#     str2 = "//td[@class=\"sm_tab_link\" and text()=\"%s\"]" % "Inpatient Cost"
#     wb = driver.find_element_by_xpath(str2)
#     action_click(wb)
#     f = InpatientCost(driver)
#     f.iterate_filter(year, customer_id)
#     loader_element = 'sm_download_cssload_loader_wrap'
#     WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
#     nav_back()

def verify_Cohort(service_year,customer_id):
    str2 = '//td[@class="sm_tab_link" and text()="Cohort Analyzer"]'
    try:
        wb = driver.find_element_by_xpath(str2)
        action_click(wb)
        print("clicked on Cohort analyzer ")
    except NoSuchElementException as e :
        print("Cohort not found ",customer_id)
        return e
        pass
    service_year=[service_year]

    f = CohortAnalyzer(driver)
    try:
        f.iterate_filter(service_year, customer_id)
        loader_element = 'sm_download_cssload_loader_wrap'
        WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
        nav_back()
    except TimeoutException as e2:
        close_button_when_loading = driver.find_element_by_xpath(
            config.get("runner", "close_button_when_loading_xpath"))
        print("Time out exception for ", customer_id, "Cohort Analyzer")
        logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in" + "Cohort Analyzer" + "\n")
        action_click(close_button_when_loading)
        nav_back()
        pass
    # except (
    #         WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
    #         StaleElementReferenceException) as e:
    #     print(e)
    #     print("Exception occurred in clicking on  " + "Cohort Analyzer")
    #     logger.error(str(e) + str(customer_id) + "Cohort Analyzer" + "\n")
    #     nav_back()
    #     pass
    # updated sm_tab_link
    # str2 = "//td[@class=\"sm_tab_link\" and text()=\"%s\"]" % "Cohort Analyzer Summary"
    #str2 = "//td[@class=\"sm_tab_link\" and text()=\"%s\"]" % "Cohort Analyzer Summary"
    str23 = '//td[@class="sm_tab_link" and text()="Cohort Analyzer Summary"]'
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, str2)))
        wb2 = driver.find_element_by_xpath(str23)
        action_click(wb2)
    except NoSuchElementException:
        print("Cohort Summary not found ", customer_id)
        return 0
        pass
    #wb.click()

    f = CohortAnalyzerSummary(driver)
    try:
        f.iterate_filter(service_year, customer_id)
        loader_element = 'sm_download_cssload_loader_wrap'
        WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
        nav_back()
    except TimeoutException as e2:
        close_button_when_loading = driver.find_element_by_xpath(
            config.get("runner", "close_button_when_loading_xpath"))
        print("Time out exception for ", customer_id, "Cohort Analyzer")
        logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in" + "Cohort Analyzer Summary" + "\n")
        action_click(close_button_when_loading)
        nav_back()
        pass
    # except (
    #         WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
    #         StaleElementReferenceException) as e:
    #     print(e)
    #     print("Exception occurred in Utilization  " + "Cohort Analyzer")
    #     logger.error(str(e) + str(customer_id) + "Cohort Analyzer Summary" + "\n")
    #     nav_back()
    #     pass


def verify_quality(year, customer_id):
    loader_element = 'sm_download_cssload_loader_wrap'
    WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
    str2 = "//td[@class=\"sm_tab_link\" and text()=\"%s\"]" % "Quality Overview"
    worksheet = driver.find_element_by_xpath(str2)
    worksheet.click()
    f = QualityOverview(driver)
    year = [year]
    try:
        f.iterate_filter(year, customer_id)
        loader_element = 'sm_download_cssload_loader_wrap'
        WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
        nav_back()
    except TimeoutException as e2:
        close_button_when_loading = driver.find_element_by_xpath(
            config.get("runner", "close_button_when_loading_xpath"))
        print("Time out exception for ", customer_id, " Quality Overview ")
        logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in Quality Overview \n")

        action_click(close_button_when_loading)
        nav_back()
        pass
    except (TimeoutException, WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
            StaleElementReferenceException) as e:
        print(e)
        print("Exception occurred in Quality Overview")
        logger.error(str(e) + str(customer_id) + " Exception occurred in Quality Overview ")
        nav_back()
        pass
    except NoSuchElementException as e:
        print("Exception occurred in Quality Overview")
        logger.error(str(e) + str(customer_id) + " Exception occurred in Quality Overview ")
        nav_back()
        pass


def verify_Commercial_risk(year, customer_id):
    worksheets = driver.find_elements_by_xpath("//tr[@workbook_title=\"Commercial Risk\"]")
    year = [year]
    for i in range(1, len(worksheets) + 1):
        worksheet_xpath = "//tr[@workbook_title=\"Commercial Risk\"][%s]" % i
        worksheet = driver.find_element_by_xpath(worksheet_xpath)
        if worksheet.get_attribute("worksheet_title") == "Risk Overview":
            print("Commercial" + worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = CommercialRiskOverview(driver)
            try:
                f.iterate_filter(year, customer_id)
                loader_element = 'sm_download_cssload_loader_wrap'
                WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
                nav_back()
            except TimeoutException as e2:
                close_button_when_loading = driver.find_element_by_xpath(
                    config.get("runner", "close_button_when_loading_xpath"))
                print("Time out exception for ", customer_id, " commercial Risk OVerview ")
                logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in Commercial Risk Overview \n")
                action_click(close_button_when_loading)
                nav_back()
                pass
            except (
                    WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
                    StaleElementReferenceException) as e:
                print(e)
                print("Exception occurred in Commercial " + worksheet.get_attribute("worksheet_title"))
                logger.error(str(e) + str(customer_id) + " Exception occurred in Commercial Risk Overview \n")
                nav_back()
                pass
        elif worksheet.get_attribute("worksheet_title") == "Coding Discontinuation":
            print("Commercial" + worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = CommercialCodingDiscontinuation(driver)
            try:
                f.iterate_filter(year, customer_id)
                loader_element = 'sm_download_cssload_loader_wrap'
                WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
                nav_back()
            except TimeoutException as e2:
                close_button_when_loading = driver.find_element_by_xpath(
                    config.get("runner", "close_button_when_loading_xpath"))
                print("Time out exception for ", customer_id, " commercial Coding disconitnuation")
                logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in " + worksheet.get_attribute(
                    "worksheet_title"), "\n")
                action_click(close_button_when_loading)
                nav_back()
                pass
            except (
                    WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
                    StaleElementReferenceException) as e:
                print(e)
                print("Exception occurred in Commercial " + worksheet.get_attribute("worksheet_title"))
                logger.error(str(e) + str(customer_id) + " Timeout Exception occurred in " + worksheet.get_attribute(
                    "worksheet_title"), "\n")
                nav_back()
                pass


def verify_medicare_risk(year, customer_id):
    worksheets = driver.find_elements_by_xpath("//tr[@workbook_title=\"Medicare Risk\"]")
    year = [year]
    for i in range(1, len(worksheets) + 1):
        worksheet_xpath = "//tr[@workbook_title=\"Medicare Risk\"][%s]" % i
        worksheet = driver.find_element_by_xpath(worksheet_xpath)
        if worksheet.get_attribute("worksheet_title") == "Risk Overview":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = MedicareRiskOverview(driver)
            try:
                f.iterate_filter(year, customer_id)
                nav_back()
            except TimeoutException as e2:
                close_button_when_loading = driver.find_element_by_xpath(
                    config.get("runner", "close_button_when_loading_xpath"))
                print("Time out exception for ", customer_id, worksheet.get_attribute("worksheet_title"))
                logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in " + worksheet.get_attribute(
                    "worksheet_title"), "\n")
                action_click(close_button_when_loading)
                nav_back()
                pass
            except (WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
                    StaleElementReferenceException) as e:
                print(e)
                print("Exception occurred in Risk overview")
                logger.error(str(e) + str(customer_id) + " Exception occurred in Medicare Risk OVerview ")
                nav_back()
                pass
        elif worksheet.get_attribute("worksheet_title") == "Suspect Analytics":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = SuspectAnalytics(driver)
            try:
                f.iterate_filter(year, customer_id)
                nav_back()
            except TimeoutException as e2:
                close_button_when_loading = driver.find_element_by_xpath(
                    config.get("runner", "close_button_when_loading_xpath"))
                print("Time out exception for ", customer_id, worksheet.get_attribute("worksheet_title"))
                logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in " +
                             worksheet.get_attribute("worksheet_title"), "\n")
                action_click(close_button_when_loading)
                nav_back()
                pass
            except (WebDriverException, ElementNotInteractableException,
                    ElementClickInterceptedException,
                    StaleElementReferenceException) as e:
                print(e)
                print("Exception occurred in Suspect Analytics")
                logger.error(str(e) + str(customer_id) + " Exception occurred in Suspect Analytics")
                nav_back()
                pass
        elif worksheet.get_attribute("worksheet_title") == "Coding Discontinuation":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = MedicareCodingDiscontinuation(driver)
            try:
                f.iterate_filter(year, customer_id)
                nav_back()
            except TimeoutException as e2:
                close_button_when_loading = driver.find_element_by_xpath(
                    config.get("runner", "close_button_when_loading_xpath"))
                print("Time out exception for ", customer_id, worksheet.get_attribute("worksheet_title"))
                logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in " +
                             worksheet.get_attribute("worksheet_title"), "\n")
                action_click(close_button_when_loading)
                nav_back()
                pass
            except (
                    TimeoutException, WebDriverException, ElementNotInteractableException,
                    ElementClickInterceptedException,
                    StaleElementReferenceException) as e:
                print(e)
                print("Exception occurred in Medicare Coding Discontinuation")
                logger.error(str(e) + str(customer_id) + " Exception occurred in Medicare Coding Discontinuation")
                nav_back()
                pass
        elif worksheet.get_attribute("worksheet_title") == "RAF Reconciliation":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = MedicareRAF(driver)
            try:
                f.iterate_filter(year, customer_id)
                nav_back()
            except TimeoutException as e2:
                close_button_when_loading = driver.find_element_by_xpath(
                    config.get("runner", "close_button_when_loading_xpath"))
                print("Time out exception for ", customer_id, worksheet.get_attribute("worksheet_title"))
                logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in "
                             + worksheet.get_attribute("worksheet_title"), "\n")
                action_click(close_button_when_loading)
                nav_back()
                pass
            except (WebDriverException, ElementNotInteractableException,
                    ElementClickInterceptedException,
                    StaleElementReferenceException) as e:
                print(e)
                print("Exception occurred in Medicare RAF Reconciliation")
                logger.error(str(e) + str(
                    customer_id) + " Exception occurred in Medicare RAF Reconciliation")
                nav_back()
                pass
        elif worksheet.get_attribute("worksheet_title") == "Risk by Hierarchy":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = MedicareRiskByHierarchy(driver)
            try:
                f.iterate_filter(year, customer_id)
                nav_back()
            except TimeoutException as e2:
                close_button_when_loading = driver.find_element_by_xpath(
                    config.get("runner", "close_button_when_loading_xpath"))
                print("Time out exception for ", customer_id, worksheet.get_attribute("worksheet_title"))
                logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in " +
                             worksheet.get_attribute("worksheet_title"), "\n")
                action_click(close_button_when_loading)
                nav_back()
                pass
            except (WebDriverException, ElementNotInteractableException,
                    ElementClickInterceptedException,
                    StaleElementReferenceException) as e:
                print(e)
                print("Exception occurred in Medicare Risk by Hierarchy")
                logger.error(str(e) + str(
                    customer_id) + " EException occurred in Medicare Risk by Hierarchy")
                nav_back()
                pass


def verify_medicaid_risk(year, customer_id):
    worksheets = driver.find_elements_by_xpath("//tr[@workbook_title=\"Medicaid Risk\"]")
    print(worksheets)
    for i in range(1, len(worksheets) + 1):
        worksheet_xpath = "//tr[@workbook_title=\"Medicaid Risk\"][%s]" % i
        worksheet = driver.find_element_by_xpath(worksheet_xpath)
        print(worksheet_xpath)
        if worksheet.get_attribute("worksheet_title") == "Risk Overview":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = MedicareRiskOverview(driver)
            f.iterate_filter(year, customer_id)
            nav_back()
        elif worksheet.get_attribute("worksheet_title") == "Suspect Analytics":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = SuspectAnalytics(driver)
            f.iterate_filter(year, customer_id)
            nav_back()
        elif worksheet.get_attribute("worksheet_title") == "Coding Discontinuation":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = MedicareCodingDiscontinuation(driver)
            f.iterate_filter(year, customer_id)
            nav_back()
        elif worksheet.get_attribute("worksheet_title") == "RAF Reconciliation":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = MedicareRAF(driver)
            f.iterate_filter(year, customer_id)
            nav_back()
        elif worksheet.get_attribute("worksheet_title") == "Risk by Hierarchy":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = MedicareRiskByHierarchy(driver)
            f.iterate_filter(year, customer_id)
            nav_back()


def verify_utilization(year, customer_id):
    worksheets = driver.find_elements_by_xpath("//tr[@workbook_title=\"Utilization\"]")
    year = [year]
    for i in range(1, len(worksheets) + 1):
        worksheet_xpath = "//tr[@workbook_title=\"Utilization\"][%s]" % i
        worksheet = driver.find_element_by_xpath(worksheet_xpath)
        if worksheet.get_attribute("worksheet_title") == "Total Cost":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = TotalCost(driver)
            try:
                f.iterate_filter(year, customer_id)
                loader_element = 'sm_download_cssload_loader_wrap'
                WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
                nav_back()
            except TimeoutException as e2:
                close_button_when_loading = driver.find_element_by_xpath(
                    config.get("runner", "close_button_when_loading_xpath"))
                print("Time out exception for ", customer_id, worksheet.get_attribute("worksheet_title"))
                logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in" + worksheet.get_attribute(
                    "worksheet_title") + "\n")
                action_click(close_button_when_loading)
                nav_back()
                pass
            except (
                    WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
                    StaleElementReferenceException) as e:
                print(e)
                print("Exception occurred in Utilization  " + worksheet.get_attribute("worksheet_title"))
                logger.error(str(e) + str(customer_id) + worksheet.get_attribute("worksheet_title") + "\n")
                nav_back()
                pass

        #elif worksheet.get_attribute("worksheet_title") == "Inpatient Cost":
        #    print(worksheet.get_attribute("worksheet_title"))
        #    worksheet.click()
        #    f = InpatientCost(driver)
        #    try:
        #        f.iterate_filter(year, customer_id)
        #        loader_element = 'sm_download_cssload_loader_wrap'
        #        WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
        #        nav_back()
        #    except TimeoutException as e2:
        #        close_button_when_loading = driver.find_element_by_xpath(
        #            config.get("runner", "close_button_when_loading_xpath"))
        #        print("Time out exception for ", customer_id, worksheet.get_attribute("worksheet_title"))
        #        logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in" + worksheet.get_attribute(
        #            "worksheet_title") + "\n")
        #        action_click(close_button_when_loading)
        #        nav_back()
        #        pass
        #    except (
        #            WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
        #            StaleElementReferenceException) as e:
        #        print(e)
        #        print("Exception occurred in Utilization  " + worksheet.get_attribute("worksheet_title"))
        #        logger.error(str(e) + str(customer_id) + worksheet.get_attribute("worksheet_title") + "\n")
        #        nav_back()
        #        pass
        elif worksheet.get_attribute("worksheet_title") == "Inpatient Cost By Dx":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = InpatientCostByDX(driver)
            try:
                f.iterate_filter(year, customer_id)
                loader_element = 'sm_download_cssload_loader_wrap'
                WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
                nav_back()
            except TimeoutException as e2:
                close_button_when_loading = driver.find_element_by_xpath(
                    config.get("runner", "close_button_when_loading_xpath"))
                print("Time out exception for ", customer_id, worksheet.get_attribute("worksheet_title"))
                logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in" + worksheet.get_attribute(
                    "worksheet_title") + "\n")
                action_click(close_button_when_loading)
                nav_back()
                pass
            except (
                    WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
                    StaleElementReferenceException) as e:
                print(e)
                print("Exception occurred in Utilization  " + worksheet.get_attribute("worksheet_title"))
                logger.error(str(e) + str(customer_id) + worksheet.get_attribute("worksheet_title") + "\n")
                nav_back()
                pass
        #elif worksheet.get_attribute("worksheet_title") == "Total Cost Trends":
        #    print(worksheet.get_attribute("worksheet_title"))
        #    worksheet.click()
        #    f = TotalCostTrends(driver)
        #    try:
        #        f.iterate_filter(year, customer_id)
        #        loader_element = 'sm_download_cssload_loader_wrap'
        #        WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
        #        nav_back()
        #    except TimeoutException as e2:
        #        close_button_when_loading = driver.find_element_by_xpath(
        #            config.get("runner", "close_button_when_loading_xpath"))
        #        print("Time out exception for ", customer_id, worksheet.get_attribute("worksheet_title"))
        #        logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in" + worksheet.get_attribute(
        #            "worksheet_title") + "\n")
        #        action_click(close_button_when_loading)
        #        nav_back()
        #        pass
        #    except (
        #            WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
        #            StaleElementReferenceException) as e:
        #        print(e)
        #        print("Exception occurred in Utilization  " + worksheet.get_attribute("worksheet_title"))
        #        logger.error(str(e) + str(customer_id) + worksheet.get_attribute("worksheet_title") + "\n")
        #        nav_back()
        #        pass
        #elif worksheet.get_attribute("worksheet_title") == "ED Cost Trends":
        #    print(worksheet.get_attribute("worksheet_title"))
        #    worksheet.click()
        #    f = EDCostTrends(driver)
        #    try:
        #        f.iterate_filter(year, customer_id)
        #        loader_element = 'sm_download_cssload_loader_wrap'
        #        WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
        #        nav_back()
        #    except TimeoutException as e2:
        #        close_button_when_loading = driver.find_element_by_xpath(
        #            config.get("runner", "close_button_when_loading_xpath"))
        #        print("Time out exception for ", customer_id, worksheet.get_attribute("worksheet_title"))
        #        logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in" + worksheet.get_attribute(
        #            "worksheet_title") + "\n")
        #        action_click(close_button_when_loading)
        #        nav_back()
        #        pass
        #    except (
        #            WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
        #            StaleElementReferenceException) as e:
        #        print(e)
        #        print("Exception occurred in Utilization  " + worksheet.get_attribute("worksheet_title"))
        #        logger.error(str(e) + str(customer_id) + worksheet.get_attribute("worksheet_title") + "\n")
        #        nav_back()
        #        pass
        # elif worksheet.get_attribute("worksheet_title") == "Professional Cost":
        #     print(worksheet.get_attribute("worksheet_title"))
        #     worksheet.click()
        #     f = ProfessionalCost(driver)
        #     f.iterate_filter(year, customer_id)
        #     loader_element = 'sm_download_cssload_loader_wrap'
        #     WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
        #     nav_back()
        elif worksheet.get_attribute("worksheet_title") == "Inpatient Cost By HCC":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = InpatientCostByHCC(driver)
            try:
                f.iterate_filter(year, customer_id)
                loader_element = 'sm_download_cssload_loader_wrap'
                WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
                nav_back()
            except TimeoutException as e2:
                close_button_when_loading = driver.find_element_by_xpath(
                    config.get("runner", "close_button_when_loading_xpath"))
                print("Time out exception for ", customer_id, worksheet.get_attribute("worksheet_title"))
                logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in" + worksheet.get_attribute(
                    "worksheet_title") + "\n")
                action_click(close_button_when_loading)
                nav_back()
                pass
            except (
                    WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
                    StaleElementReferenceException) as e:
                print(e)
                print("Exception occurred in Utilization  " + worksheet.get_attribute("worksheet_title"))
                logger.error(str(e) + str(customer_id) + worksheet.get_attribute("worksheet_title") + "\n")
                nav_back()
                pass
        elif worksheet.get_attribute("worksheet_title") == "ED Cost":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = EDCost(driver)
            try:
                f.iterate_filter(year, customer_id)
                loader_element = 'sm_download_cssload_loader_wrap'
                WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
                nav_back()
            except TimeoutException as e2:
                close_button_when_loading = driver.find_element_by_xpath(
                    config.get("runner", "close_button_when_loading_xpath"))
                print("Time out exception for ", customer_id, worksheet.get_attribute("worksheet_title"))
                logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in" + worksheet.get_attribute(
                    "worksheet_title") + "\n")
                action_click(close_button_when_loading)
                nav_back()
                pass
            except (
                    WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
                    StaleElementReferenceException) as e:
                print(e)
                print("Exception occurred in Utilization  " + worksheet.get_attribute("worksheet_title"))
                logger.error(str(e) + str(customer_id) + worksheet.get_attribute("worksheet_title") + "\n")
                nav_back()
                pass
        elif worksheet.get_attribute("worksheet_title") == "Pharmacy Cost":
            print(worksheet.get_attribute("worksheet_title"))
            worksheet.click()
            f = PharmacyCost(driver)
            try:
                f.iterate_filter(year, customer_id)
                loader_element = 'sm_download_cssload_loader_wrap'
                WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
                nav_back()
            except TimeoutException as e2:
                close_button_when_loading = driver.find_element_by_xpath(
                    config.get("runner", "close_button_when_loading_xpath"))
                print("Time out exception for ", customer_id, worksheet.get_attribute("worksheet_title"))
                logger.error(str(e2) + str(customer_id) + " Timeout Exception occurred in" + worksheet.get_attribute(
                    "worksheet_title") + "\n")
                action_click(close_button_when_loading)
                nav_back()
                pass
            except (
                    WebDriverException, ElementNotInteractableException, ElementClickInterceptedException,
                    StaleElementReferenceException) as e:
                print(e)
                print("Exception occurred in Utilization  " + worksheet.get_attribute("worksheet_title"))
                logger.error(str(e) + str(customer_id) + worksheet.get_attribute("worksheet_title") + "\n")
                nav_back()
                pass


def get_customer_ids():
    # change for user
    db_path = config.get("runner", "dbpath")
    folder_path = ''
    conn = create_connection(db_path)
    cur = conn.cursor()
    cur.execute("select * from suremetrics_log_entry_stage")
    rows = cur.fetchall()
    customer_list = []
    customer_list2 = []
    for row in rows:
        customer_list.append(row[0])
        customer_list2.append(row[1])
    return tuple(zip(customer_list, customer_list2))


def open_analytics_page(customer_id):
    customer_list_url = []
    sm_customer_id = customer_id
    session_var = 'app_id=analytics&custId=' + str(sm_customer_id) + '&payerId=' + str(
        sm_customer_id) + '&orgId=' + str(sm_customer_id)
    encoded_string = base64.b64encode(session_var.encode('utf-8'))
    customer_list_url.append(encoded_string)
    for idx, val in enumerate(customer_list_url):
        driver.get("https://www.cozeva.com/analytics/?session=" + val.decode('utf-8'))


def open_analytics_risk_page(customer_id):
    customer_list_url = []
    sm_customer_id = customer_id
    session_var = 'app_id=analytics&custId=' + str(sm_customer_id) + '&payerId=' + str(
        sm_customer_id) + '&orgId=' + str(sm_customer_id)
    encoded_string = base64.b64encode(session_var.encode('utf-8'))
    customer_list_url.append(encoded_string)
    for idx, val in enumerate(customer_list_url):
        driver.get("https://www.cozeva.com/analytics/risk?session=" + val.decode('utf-8'))


def open_analytics_utilization_page(customer_id):
    customer_list_url = []
    sm_customer_id = customer_id
    session_var = 'app_id=analytics&custId=' + str(sm_customer_id) + '&payerId=' + str(
        sm_customer_id) + '&orgId=' + str(sm_customer_id)
    encoded_string = base64.b64encode(session_var.encode('utf-8'))
    customer_list_url.append(encoded_string)
    for idx, val in enumerate(customer_list_url):
        driver.get("https://www.cozeva.com/analytics/utilization?session=" + val.decode('utf-8'))


def get_error(db_loc):
    try:
        conn = sqlite3.connect(db_loc)
    except Error as e:
        print(e)
    cur = conn.cursor().execute("Select * from analytics_nodata_found")
    rows = cur.fetchall()
    customer_id = []
    workbook = []
    year = []
    drilldown = []
    lob = []
    f = open("No-Data-Found.csv", "w")
    f.write("Serial,Customer,Workbook,Year,LOB,Drill Down\n")
    i = 0
    for row in rows:
        i = i + 1
        customer_id.append(row[0])
        workbook.append(row[1])
        year.append(row[2])
        drilldown.append(row[3])
        lob.append(row[4])
        f.write("{},{}, {},{},{},{}\n".format(i, row[0], row[1], row[2], row[4], row[3]))


# Logger Settings
source_directory = os.getcwd()
db_path = config.get("runner", "dbpath")
config.set('runner', 'dbpath', source_directory + "\\assets\\suremetrics_log.db")
db_path = config.get("runner", "dbpath")
print("After change= " + db_path)
config.write(config_path.open("w"))
config_path = Path("locator-config.properties")
config.read(config_path)

from QualityOverview import QualityOverview
from TotalCost import TotalCost
from EDCost import EDCost
from InpatientCost import InpatientCost
from MedicareRiskOverview import MedicareRiskOverview
from SuspectAnalytics import SuspectAnalytics
from MedicareCodingDiscontinuation import MedicareCodingDiscontinuation
from PharmacyCost import PharmacyCost
from MedicareRAF import MedicareRAF
from MedicareRiskByHierarchy import MedicareRiskByHierarchy
from CommercialCodingDiscontinuation import CommercialCodingDiscontinuation
from CommercialRiskOverview import CommercialRiskOverview
from InpatientCostByDX import InpatientCostByDX
from InpatientCostByHCC import InpatientCostByHCC
import TotalCostTrends
from TotalCostTrends import TotalCostTrends
from EDCostTrends import EDCostTrends
from CohortAnalyzer import CohortAnalyzer
from CohortAnalyzerSummary import CohortAnalyzerSummary
from UsageMonthlyActivity import UsageMonthlyActivity

print(source_directory)
dateandtime = date_time()
master_directory = config.get("runner", "report_directory_input")
os.chdir(master_directory)
path = makedir(dateandtime)
LOG_FORMAT = "%(levelname)s %(asctime)s - %(message)s"
logging.basicConfig(filename=path + "\\" + "Error-Log.log", level=logging.INFO, format=LOG_FORMAT, filemode='w')
logger = logging.getLogger()
logger.setLevel(logging.ERROR)
os.chdir(path)
downloaddefault = config.get("runner", "downloaddefault")
downloaddefault = source_directory + "\\" + downloaddefault
driver = setup("Chrome", downloaddefault)

makedir(downloaddefault)
# driver = setup("Chrome",downloaddefault)
# driver = setups.driver_setup()
begin_time = datetime.now()
# change for user
loc = config.get("runner", "login_file")
# add command line argument for path
# log file

login(driver, loc)
logger.info("Login successful")
truncate_table("analytics_nodata_found")

location = source_directory + "\\" + config.get("runner", "config_file")

wb = xlrd.load_workbook(location)
sheet = wb.active
# sheet.cell_value(0, 0)

num_of_rows = sheet.max_row
print("No of instances ", num_of_rows)
logger.info("No of instances " + str(num_of_rows))
starting_row = int(config.get("runner", "starting_row"))
till_row = int(config.get("runner", "till_row"))

report_directory = os.getcwd()

for i in range(2, num_of_rows + 1):
    parameter_list = sheet[i]
    print(parameter_list)
    customer_id = str(int(parameter_list[0].value)).replace(" ", "")
    customer_name = str(parameter_list[1].value).replace(" ", "")
    service_year = str(parameter_list[2].value).strip()
    if "Q" in service_year:
        service_year = str(parameter_list[2].value).strip()
    else:
        service_year = str(int(parameter_list[2].value)).replace(" ", "")
    medicare = str(parameter_list[3].value).replace(" ", "")
    commercial = str(parameter_list[4].value).replace(" ", "")
    utilization = str(parameter_list[5].value).replace(" ", "")
    usage = str(parameter_list[6].value).replace(" ", "")
    open_analytics_page(customer_id)
    logger.info("Opened Analytics Page")
    make_directory(customer_id)
    logger.info("Made Directory of Customer ")
    make_directory(customer_id)
    verify_quality(service_year, customer_id)
    if (medicare == 'Y'):
        verify_medicare_risk(service_year, customer_id)
    if (commercial == 'Y'):
        verify_Commercial_risk(service_year, customer_id)
    if (utilization == "Y"):
        verify_utilization(service_year, customer_id)
    verify_Cohort(service_year, customer_id)
    if (usage == "Y"):
        verify_Usage(service_year, customer_id)

get_error(config.get("runner", "dbpath"))
os.chdir("C://")
driver.quit()
notification_title = "Analytics Execution Complete "

notification_message = "Please check report here" + "C:\\VerificationReports\\ExportReport\\"

notification.notify(
            title=notification_title,
            message=notification_message,
            app_icon=None,
            timeout=20,
            toast=False
        )

print("Execution time ", datetime.now() - begin_time)
config_path = Path(source_directory + "\\" + "locator-config.properties")
config.set('runner', 'dbpath', "assets\\suremetrics_log.db")
config.write(config_path.open("w"))