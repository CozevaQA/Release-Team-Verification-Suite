from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, \
    ElementClickInterceptedException, UnexpectedAlertPresentException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import sqlite3
from sqlite3 import Error
import os
from os import path
import shutil
import configparser
import time
from DownloadWorkbookData import download_workbook_data,copy_paste_file
# contains year ,lob, drill down


config = configparser.RawConfigParser()
config.read("locator-config.properties")



def create_connection(db_file):  # creating connection
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except Error as e:
        print(e)

    return None

def wait_to_load(driver):
    loader=config.get("CohortAnalyzer-Prod","loader_element")
    WebDriverWait(driver,200).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader)))

def wait_to_load_filter(driver):
    loader=config.get("CohortAnalyzer-Prod","loader_element_filter")
    try :
        WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader)))
    except UnexpectedAlertPresentException:
        print("Unknown Error Occurred while loading page ")


class CohortAnalyzer:

    def __init__(self, driver):
        db_path = config.get("runner","dbpath")
        folder_path = ''
        self.conn = create_connection(db_path)
        self.cur = self.conn.cursor()
        self.driver = driver
        self.loader_element = config.get("CohortAnalyzer-Prod","loader_element")
        self.selected_value_year_xpath = config.get("CohortAnalyzer-Prod","selected_value_year_xpath")
        self.service_year_xpath = config.get("CohortAnalyzer-Prod","service_year_xpath")
        self.lob_xpath = config.get("CohortAnalyzer-Prod","lob_xpath")
        self.apply_filter_xpath = config.get("CohortAnalyzer-Prod","apply_filter_xpath")
        self.loader_panel_class = config.get("CohortAnalyzer-Prod","loader_panel_class")
        self.lob_outer_elements_xpath = config.get("CohortAnalyzer-Prod","lob_outer_elements_xpath")
        self.insert_chart_xpath=config.get("CohortAnalyzer-Prod","insert_chart_xpath")
        self.filtermodal_xpath=config.get("CohortAnalyzer-Prod","filter_xpath")


    def check_exists_byclass(self, classname):
        try:
            self.driver.find_element_by_class_name(classname)
        except NoSuchElementException:
            return False
        return True

    def action_click(self,element):
        try:
            #element.click()
            ActionChains(self.driver).move_to_element(element).click(element).perform()

        except (ElementNotInteractableException,ElementClickInterceptedException):
            self.driver.execute_script("arguments[0].click();", element)

    def makedir(self, customer):
        path1 = str(customer) + "/" + str(customer) + "-Cohort Analyzer"
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

    def iterate_filter(self, year, customer):
        wbpath = self.makedir(customer)
        #report_directory=os.path.join(report_directory,wbpath)
        wait_to_load(self.driver)
        self.action_click(self.driver.find_element_by_xpath(self.insert_chart_xpath))
        lob_outer_elements = self.driver.find_elements_by_xpath(self.lob_outer_elements_xpath)
        count_outer = len(lob_outer_elements)
        #print(count_outer)
        loop_list = []
        count = 1
        loop_list.append(count_outer)
        lob_values = []
        lob_names = []
        for j in range(1, count_outer + 1):
            lob_inner_elements_xpath = self.lob_outer_elements_xpath + "[" + str(j) + "]//option"
            lob_inner_elements = self.driver.find_elements_by_xpath(lob_inner_elements_xpath)
            for k in range(1, len(lob_inner_elements) + 1):
                lob_inner_element_xpath = lob_inner_elements_xpath + "[" + str(k) + "]"
                #print(lob_inner_element_xpath)
                lob_name = self.driver.find_element_by_xpath(lob_inner_element_xpath).get_attribute("innerHTML")
                lob_value = self.driver.find_element_by_xpath(lob_inner_element_xpath).get_attribute("value")
                lob_values.append(lob_value)
                lob_names.append(lob_name)

            first_time = 1
            num_of_child = len(lob_inner_elements)
            loop_list.append(num_of_child)

        #print(loop_list)
        print(lob_values)
        for y in year:
            wait_to_load(self.driver)
            selected_value = self.driver.find_element_by_xpath(self.selected_value_year_xpath).text
            if (customer == str('200')):
                service_year_for_filter = y[0:4]
            else:
                service_year_for_filter = y
            selected_value = self.driver.find_element_by_xpath(self.selected_value_year_xpath).text
            if int(selected_value) != service_year_for_filter:

                service_year = self.driver.find_element_by_xpath(self.service_year_xpath)

                ActionChains(self.driver).move_to_element(service_year).click(service_year).perform()
                year_selector = "//input[@type=\"radio\" and @value=\"%s\"]" % str(service_year_for_filter)
                try:
                    ele = self.driver.find_element_by_xpath(year_selector)
                    ele.location_once_scrolled_into_view
                    self.action_click(ele)
                    #ActionChains(self.driver).move_to_element(service_year).click(service_year).perform()
                    wait_to_load_filter(self.driver)
                except NoSuchElementException:
                    with self.conn:
                        self.cur.execute(
                            'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                            (int(customer), "Cohort Analyzer", int(y), str("Does Not Exists"), str("Does not exists"),))

                    print(service_year_for_filter, "does not exist ")
                    break

            print(y)

            j=0
            count=1
            for i in range(1,len(loop_list)):
                count=1
                first_entry=1
                for x in range(1,loop_list[i]+1):
                    if(count>1):
                        lob = self.driver.find_element_by_xpath(self.lob_xpath)
                        self.driver.execute_script("arguments[0].click();", lob)
                        #print("Clicked on LOB select dropdown")
                    count=count+1
                    if(first_entry==1):
                        value_in_input = str(lob_values[j-1])
                        # print(value_in_input)
                        lobtobeclicked_xpath = "//input[@type=\"checkbox\" and @value=\"%s\"]//following-sibling::span" % (value_in_input)
                        # print(lobtobeclicked_xpath)
                        lobtobeclicked = self.driver.find_element_by_xpath(lobtobeclicked_xpath)
                        self.action_click(lobtobeclicked)
                        #print("First Value is unchecked")


                        first_entry=0

                    value_in_input=str(lob_values[j])
                    name_of_lob = str(lob_names[j]).replace(" ", "")
                    print(name_of_lob)
                    lobtobeclicked_xpath="//input[@type=\"checkbox\" and @value=\"%s\"]//following-sibling::span" %(value_in_input)
                    #print(lobtobeclicked_xpath)

                    lobtobeclicked=self.driver.find_element_by_xpath(lobtobeclicked_xpath)
                    self.action_click(lobtobeclicked)
                    print("LOB to be clicked is selected ")
                    time.sleep(1)
                    #click on filter modal

                    j=j+1

                    not_closed=0
                    while(not_closed<=1):
                        try:
                            #close the open modals
                            open_modal_xpath='(//div[@class="btn-group open"])[1]'
                            try:
                                open_modal = self.driver.find_element_by_xpath(open_modal_xpath)
                                self.driver.execute_script("arguments[0].setAttribute('class',arguments[1])",
                                                           open_modal, 'btn-group close')
                            except NoSuchElementException:
                                pass

                            #enable apply button
                            disabled_apply_xpath = '//a[@id="sm_dashboard_filter_apply"]'
                            apply_button = self.driver.find_element_by_xpath(disabled_apply_xpath)
                            self.driver.execute_script("arguments[0].setAttribute('class',arguments[1])", apply_button,
                                                       'pull-right sm_enabled')

                            # drop_down_xpath='(//select[@name="lob_metric"]//following-sibling::div//child::button)[1]'
                            # drop_down = self.driver.find_element_by_xpath(drop_down_xpath)
                            # #ActionChains(self.driver).move_to_element(drop_down).click(drop_down).perform()
                            # #self.driver.execute_script('arguments[0].setAttribute("aria-expanded", "false"); ',drop_down)
                            #
                            #
                            # # drop_down = self.driver.find_element_by_xpath(drop_down_xpath)
                            # # drop_down_state=drop_down.get_attribute("aria-expanded")
                            # # if(drop_down_state=="false"):
                            # #     not_closed=False
                            # self.driver.execute_script("arguments[0].click();", drop_down)
                            # print("Clicked on LOB select dropdown to close")
                            not_closed=not_closed+1
                        except (ElementNotInteractableException, NoSuchElementException) as e:
                            print(e)



                    #ActionChains(self.driver).move_to_element(self.driver.find_element_by_xpath(self.apply_filter_xpath)).click(self.driver.find_element_by_xpath(self.apply_filter_xpath)).perform()
                    self.action_click(self.driver.find_element_by_xpath(self.apply_filter_xpath))
                    wait_to_load(self.driver)
                    #wait to load select element

                    drop_down_xpath = '(//select[@name="lob_metric"]//following-sibling::div//child::button)[1]'
                    drop_down = self.driver.find_element_by_xpath(drop_down_xpath)

                    try:
                        element = WebDriverWait(self.driver, 300).until(
                            EC.element_to_be_clickable((By.XPATH,drop_down_xpath)))
                    except ElementNotInteractableException as e:
                        print(e)

                    if self.check_exists_byclass("nodata"):
                        print("No data found in {} for lob {}".format(y, value_in_input))
                        try:
                            with self.conn:
                                self.cur.execute(
                                    'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                    (int(customer), "Cohort Analyzer", int(y), str("Overview"), str(value_in_input),))
                                # print 'Data saved from exception'
                                ss_name = str(wbpath) + "/" + str(y) + str(name_of_lob.replace(" ", "")) + ".png"  # naming convention is year-lob-drill.png
                                print(ss_name)
                                self.driver.save_screenshot(ss_name)
                        except sqlite3.IntegrityError:
                            pass

                    else:
                        ss_name = str(wbpath) + "/" + str(y) + str(
                            name_of_lob.replace(" ", "")) + ".png"  # naming convention is year-lob-drill.png
                        print(ss_name)
                        self.driver.save_screenshot(ss_name)
                        print("data found")
                        # download_workbook_data(self.driver, downloaddefault,wbpath, customer, "Cohort",y, value_in_input[0:4])
                        # print("File downloaded at ",downloaddefault," to be copied to ",wbpath)
                        # wait_to_load(self.driver)

            # copy_paste_file(downloaddefault,os.path.join(report_directory,wbpath))
            # Run a loop i from list[1] to list[last]
            # Run a loop from 1 to i
            # lob = self.driver.find_element_by_xpath(self.lob_xpath)
            # self.driver.execute_script("arguments[0].click();", lob)


                #
    #
        # for k in range(1, len(lob_inner_elements) + 1):
                #     if(first_time>1):
                #         self.driver.execute_script("arguments[0].click();", lob)
                #     lob_inner_element_xpath = lob_inner_elements_xpath + "[" + str(k) + "]"
                #     #print(lob_inner_element_xpath)
                #     lob_inner_element = self.driver.find_element_by_xpath(lob_inner_element_xpath)
                #     self.driver.execute_script("arguments[0].click();", lob_inner_element)
                #     val = lob_inner_element.text
                #     print(val)










# create a class to open Cohort Analyzer session in stage from mysql
# incorporate this class and see how it works
# create another class for reporting mail
