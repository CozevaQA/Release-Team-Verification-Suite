from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, \
    ElementClickInterceptedException, UnexpectedAlertPresentException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
import runner
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import sqlite3
from sqlite3 import Error
import os
from os import path
import shutil
import configparser

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

def wait_to_load_filter(driver):
    loader=config.get("CohortAnalyzer-Prod","loader_element_filter")
    try :
        WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader)))
    except UnexpectedAlertPresentException:
        print("Unknown Error Occurred while loading page ")


def wait_to_load(driver):
    loader=config.get("CohortAnalyzerSummary-Prod","loader_element")
    WebDriverWait(driver,100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader)))


class CohortAnalyzerSummary:

    def __init__(self, driver):
        db_path = config.get("runner","dbpath")
        folder_path = ''
        self.conn = create_connection(db_path)
        self.cur = self.conn.cursor()
        self.driver = driver
        self.loader_element = config.get("CohortAnalyzerSummary-Prod","loader_element")
        self.selected_value_year_xpath = config.get("CohortAnalyzerSummary-Prod","selected_value_year_xpath")
        self.service_year_xpath = config.get("CohortAnalyzerSummary-Prod","service_year_xpath")
        self.lob_xpath = config.get("CohortAnalyzerSummary-Prod","lob_xpath")
        self.apply_filter_xpath = config.get("CohortAnalyzerSummary-Prod","apply_filter_xpath")
        self.loader_panel_class = config.get("CohortAnalyzerSummary-Prod","loader_panel_class")
        self.lob_outer_elements_xpath = config.get("CohortAnalyzerSummary-Prod","lob_outer_elements_xpath")
        self.insert_chart_xpath=config.get("CohortAnalyzerSummary-Prod","insert_chart_xpath")
        self.member_all_count_xpath=config.get("CohortAnalyzerSummary-Prod","member_all_count_xpath")
        self.member_cohort_count_xpath=config.get("CohortAnalyzerSummary-Prod","member_cohort_count_xpath")

    def check_exists_byclass(self, classname):
        try:
            self.driver.find_element_by_class_name(classname)
        except NoSuchElementException:
            return False
        return True

    def action_click(self,element):
        try:
            element.click()
        except (ElementNotInteractableException,ElementClickInterceptedException):
            self.driver.execute_script("arguments[0].click();", element)

    def makedir(self, customer):
        path1 = str(customer) + "/" + str(customer) + "-Cohort Analyzer Summary"
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
        wait_to_load(self.driver)
        runner.remove_chat_dashboard()
        #self.driver.find_element_by_xpath(self.insert_chart_xpath).click()
        lob_outer_elements = self.driver.find_elements_by_xpath(self.lob_outer_elements_xpath)
        count_outer = len(lob_outer_elements)
        print(count_outer)
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
                lob_value = self.driver.find_element_by_xpath(lob_inner_element_xpath).get_attribute("value")
                lob_values.append(lob_value)
                lob_name = self.driver.find_element_by_xpath(lob_inner_element_xpath).get_attribute("innerHTML")
                lob_names.append(lob_name)
            first_time = 1
            num_of_child = len(lob_inner_elements)
            loop_list.append(num_of_child)

        #print(loop_list)
        #print(lob_values)
        for y in year:
            wait_to_load(self.driver)
            selected_value = self.driver.find_element_by_xpath(self.selected_value_year_xpath).text
            if int(selected_value) != y:

                service_year = self.driver.find_element_by_xpath(self.service_year_xpath)

                ActionChains(self.driver).move_to_element(service_year).click(service_year).perform()
                year_selector = "//label[@class=\"radio\" and @title=\"%s\"]" % str(y)
                try:
                    ele = self.driver.find_element_by_xpath(year_selector)
                    ele.location_once_scrolled_into_view
                    self.action_click(ele)
                    wait_to_load_filter(self.driver)
                except NoSuchElementException:
                    with self.conn:
                        self.cur.execute(
                            'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                            (int(customer), "Cohort Analyzer Summary", int(y), str("Does Not Exists"), str("Does not exists"),))

                    print(y, "does not exist ")
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
                    count=count+1
                    if(first_entry==1):
                        value_in_input = str(lob_values[j-1])
                        # print(value_in_input)
                        lobtobeclicked_xpath = "//input[@type=\"checkbox\" and @value=\"%s\"]//following-sibling::span" % (value_in_input)
                        # print(lobtobeclicked_xpath)
                        lobtobeclicked = self.driver.find_element_by_xpath(lobtobeclicked_xpath)
                        self.action_click(lobtobeclicked)
                        first_entry=0

                    value_in_input=str(lob_values[j])
                    name_of_lob = str(lob_names[j]).replace(" ", "")
                    print(name_of_lob)
                    lobtobeclicked_xpath="//input[@type=\"checkbox\" and @value=\"%s\"]//following-sibling::span" %(value_in_input)
                    #print(lobtobeclicked_xpath)
                    lobtobeclicked=self.driver.find_element_by_xpath(lobtobeclicked_xpath)
                    self.action_click(lobtobeclicked)
                    j=j+1
                    not_closed = 0
                    while (not_closed <= 1):
                        try:
                            # close the open modals
                            open_modal_xpath = '(//div[@class="btn-group open"])[1]'
                            try:
                                open_modal = self.driver.find_element_by_xpath(open_modal_xpath)
                                self.driver.execute_script("arguments[0].setAttribute('class',arguments[1])",
                                                           open_modal, 'btn-group close')
                            except NoSuchElementException:
                                print("Unable to close modal in cohort for {}".format(name_of_lob))
                                pass

                            # enable apply button
                            try:
                                disabled_apply_xpath = '//a[@id="sm_dashboard_filter_apply"]'
                                apply_button = self.driver.find_element_by_xpath(disabled_apply_xpath)
                                self.driver.execute_script("arguments[0].setAttribute('class',arguments[1])",
                                                           apply_button,
                                                           'pull-right sm_enabled')
                            except:
                                pass
                            not_closed = not_closed + 1
                        except (ElementNotInteractableException, NoSuchElementException) as e:
                            print(e)
                    self.action_click(self.driver.find_element_by_xpath(self.apply_filter_xpath))
                    wait_to_load(self.driver)
                    if self.check_exists_byclass("nodata"):
                        print("No data found in {} for lob {}".format(y, value_in_input))
                        try:
                            with self.conn:
                                self.cur.execute(
                                    'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                    (int(customer), "Cohort Analyzer Summary", int(y), str("Overview"), str(value_in_input),))
                                # print 'Data saved from exception'
                                ss_name = str(wbpath) + "/" + str(y) +str(name_of_lob.replace(" ", "")) + ".png"  # naming convention is year-lob-drill.png
                                print(ss_name)
                                self.driver.save_screenshot(ss_name)
                        except sqlite3.IntegrityError:
                            pass

                    else:
                        ss_name = str(wbpath) + "/" + str(y) + str(name_of_lob.replace(" ", "")) + ".png"  # naming convention is year-lob-drill.png
                        print(ss_name)
                        self.driver.save_screenshot(ss_name)
                        member_all_count=self.driver.find_element_by_xpath(self.member_all_count_xpath).text
                        member_cohort_count=self.driver.find_element_by_xpath(self.member_cohort_count_xpath).text
                        print("Member Count: All", member_all_count)
                        print("Member Count : Cohort",member_cohort_count)
                        if(str(member_all_count)=="0.000"):
                            print("Count all is zero")
                        if (str(member_cohort_count) == "0.000"):
                            print("Count cohort is zero")
                        print("\n")
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
