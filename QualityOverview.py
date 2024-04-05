from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, \
    ElementClickInterceptedException, StaleElementReferenceException, UnexpectedAlertPresentException
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

def wait_to_load(driver):
    loader=config.get("QualityOverview-Prod","loader_element")
    WebDriverWait(driver,300).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader)))
def wait_to_load_filter(driver):
    loader=config.get("CohortAnalyzer-Prod","loader_element_filter")
    try :
        WebDriverWait(driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader)))
    except UnexpectedAlertPresentException:
        print("Unknown Error Occurred while loading page ")

class QualityOverview:

    def __init__(self, driver):
        db_path = config.get("runner","dbpath")
        folder_path = ''
        self.conn = create_connection(db_path)
        self.cur = self.conn.cursor()
        self.driver = driver
        self.loader_element = config.get("QualityOverview-Prod","loader_element")
        self.selected_value_year_xpath = config.get("QualityOverview-Prod","selected_value_year_xpath")
        self.service_year_xpath = config.get("QualityOverview-Prod","service_year_xpath")
        self.lob_xpath = config.get("QualityOverview-Prod","lob_xpath")
        #self.lob_xpath2=config.get("QualityOverview-Prod","lob_xpath2")
        self.lob_elements_xpath = config.get("QualityOverview-Prod","lob_elements_xpath")
        self.lob_select_all_xpath = config.get("QualityOverview-Prod","lob_select_all_xpath")
        self.apply_filter_xpath = config.get("QualityOverview-Prod","apply_filter_xpath")
        self.download_xpath = config.get("QualityOverview-Prod","download_xpath")
        self.download_workbook_data_xpath = config.get("QualityOverview-Prod","download_workbook_data_xpath")
        self.download_button_xpath = config.get("QualityOverview-Prod","loader_element")
        self.loader_panel_class = config.get("QualityOverview-Prod","loader_panel_class")
        self.drilldown_elements_xpath = config.get("QualityOverview-Prod","drilldown_elements_xpath")
        self.select_all_id = config.get("QualityOverview-Prod","select_all_id")
        self.overview_xpath = config.get("QualityOverview-Prod","overview_xpath")

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
        path1 = str(customer) + "/" + str(customer) + "-Quallity Overview"
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
        #remove banner announcement

        for i in year:
            wait_to_load(self.driver)
            runner.remove_chat_dashboard()
            selected_value = self.driver.find_element_by_xpath(self.selected_value_year_xpath).text
            #print selected value

            if int(selected_value) != i:

                service_year = self.driver.find_element_by_xpath(self.service_year_xpath)

                self.action_click(service_year)
                year_selector = "//label[@class=\"radio\" and @title=\"%s\"]" % str(i)
                try:
                    ele = self.driver.find_element_by_xpath(year_selector)
                    ele.location_once_scrolled_into_view
                    self.action_click(ele)
                    wait_to_load_filter(self.driver)
                except NoSuchElementException:
                    with self.conn:
                        self.cur.execute(
                            'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                            (int(customer), "Quality Overview", int(i), str("Does not exists"), str("Does not exists"),))
                    print(i, "does not exist ")
                    break
            print(i)
            lob = self.driver.find_element_by_xpath(self.lob_xpath)
            try:
                self.action_click(lob)
            except StaleElementReferenceException :
                self.driver.implicitly_wait(10)
                self.action_click(lob)


            lob_elements = self.driver.find_elements_by_xpath(self.lob_elements_xpath)
            print("Number of LOBs ", len(lob_elements))
            count = 1
            for lob_element in lob_elements:
                wait_to_load(self.driver)
                if (count > 1):
                    self.action_click(lob)
                sel_all = self.driver.find_element_by_xpath(self.lob_select_all_xpath)
                if(len(lob_elements)==1):
                    self.action_click(sel_all)
                else:
                    try:
                        sel_all.click()
                    except (ElementNotInteractableException, ElementClickInterceptedException):

                        self.driver.execute_script("arguments[0].click();", sel_all)
                    try:
                        sel_all.click()
                    except (ElementNotInteractableException, ElementClickInterceptedException):

                        self.driver.execute_script("arguments[0].click();", sel_all)
                print(lob_element.get_attribute("value"))
                val = lob_element.get_attribute("value")
                lob_selector = "//label[@class=\"checkbox\"]//input[@value=\"%s\"]/following-sibling::span" % (val)
                # print(st)
                lob_selector_b=self.driver.find_element_by_xpath(lob_selector)
                self.driver.execute_script("arguments[0].click();", lob_selector_b)
                apply_filter=self.driver.find_element_by_xpath(self.apply_filter_xpath)
                self.action_click(apply_filter)
                self.action_click(apply_filter)
                # essential wait for loading page
                wait_to_load(self.driver)
                if self.check_exists_byclass("nodata"):
                    print("No data found in {} for lob {}".format(i, val))
                    try:
                        with self.conn:
                            self.cur.execute(
                                'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                (int(customer), "Quality Overview", int(i), str("Overview"), str(val),))
                            # print 'Data saved from exception'
                            ss_name = str(wbpath) + "/" + str(i) + str(
                                val) + "Overview" + ".png"  # naming convention is year-lob-drill.png
                            self.driver.save_screenshot(ss_name)
                    except sqlite3.IntegrityError:
                        pass
                else:
                    ss_name = str(wbpath) + "/" + str(i) + str(
                        val) + "Overview" + ".png"  # naming convention is year-lob-drill.png
                    self.driver.save_screenshot(ss_name)
                    drilldown_elements = self.driver.find_elements_by_xpath(self.drilldown_elements_xpath)
                    x = len(drilldown_elements)
                    b = 0
                    for j in range(2, x + 1):
                        # print("b value initially ,j  value initially ", b, j)
                        if b > 0:
                            continue
                        wait_to_load(self.driver)
                        select_all=self.driver.find_element_by_id(self.select_all_id)
                        self.action_click(select_all)
                        drilldown_element = "//div[@class=\"breadcrumb_dropdown\"]//child::a[%s]" % (j)
                        # print(drilldown_element)
                        ele = self.driver.find_element_by_xpath(drilldown_element)
                        self.action_click(ele)
                        drill_name = ele.get_attribute("drill_down_name")
                        print(drill_name)
                        wait_to_load(self.driver)
                        if self.check_exists_byclass("nodata"):
                            # print("No data found first , value of j is ", j)
                            print(
                                "No data found in year {} , drill down {} for lob {} ".format(i, drill_name, val))
                            try:
                                with self.conn:
                                    self.cur.execute(
                                        'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                        (int(customer), "Quality Overview", int(i), str(drill_name), str(val),))
                                    # print 'Data saved from exception'
                                    ss_name = str(wbpath) + "/" + str(i) + str(val) + str(
                                        drill_name) + ".png"  # naming convention is year-lob-drill.png
                                    self.driver.save_screenshot(ss_name)
                            except sqlite3.IntegrityError:
                                pass

                            a = j - 1
                            b = j + 1
                            while b <= x:
                                prev_drill_xpath = "//div[@class=\"breadcrumb_dropdown\"]//child::a[%s]" % (a)
                                prev_drill = self.driver.find_element_by_xpath(prev_drill_xpath)
                                self.action_click(prev_drill)
                                wait_to_load(self.driver)
                                self.driver.find_element_by_id(self.select_all_id).click()
                                next_drill_xpath = "//div[@class=\"breadcrumb_dropdown\"]//child::a[%s]" % (b)
                                next_drill = self.driver.find_element_by_xpath(next_drill_xpath)
                                self.action_click(next_drill)
                                drill_name2 = next_drill.get_attribute("drill_down_name")
                                print(drill_name2)
                                wait_to_load(self.driver)
                                if self.check_exists_byclass("nodata"):
                                    # print("No data found second , value of b is {} and j is {} is ", format(b, j))
                                    print(
                                        "No data found in year {} , drill down {} for lob {} ".format(i, drill_name2,
                                                                                                      val))
                                    try:
                                        with self.conn:
                                            self.cur.execute(
                                                'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                                (
                                                    int(customer), "Quality Overview", int(i), str(drill_name2),
                                                    str(val),))
                                            # print 'Data saved from exception'
                                            ss_name = str(wbpath) + "/" + str(i) + str(val) + str(
                                                drill_name2) + ".png"  # naming convention is year-lob-drill.png
                                            self.driver.save_screenshot(ss_name)
                                    except sqlite3.IntegrityError:
                                        pass
                                else:
                                    ss_name = str(wbpath) + "/" + str(i) + str(val) + str(
                                        drill_name2) + ".png"  # naming convention is year-lob-drill.png
                                    self.driver.save_screenshot(ss_name)
                                    print("data found")

                                b = b + 1
                                # print("checking increment ! b and j is {} , {} respectively", b, j)

                        else:
                            ss_name = str(wbpath) + "/" + str(i) + str(val) + str(
                                drill_name) + ".png"  # naming convention is year-lob-drill.png
                            self.driver.save_screenshot(ss_name)
                            print("data found")
                            j = j + 1
                        # print("j value in the end ",j)
                    loader_element = 'sm_download_cssload_loader_wrap'
                    wait_to_load(self.driver)
                    self.driver.find_element_by_xpath(self.overview_xpath).click()
                    count = count + 1

    def download_data(self):
        WebDriverWait(self.driver, 20).until(EC.visibility_of_element_located((By.XPATH, self.download_xpath)))
        download = self.driver.find_element_by_xpath(self.download_xpath)
        download.click()
        download_workbook_data = self.driver.find_element_by_xpath(self.download_workbook_data_xpath)
        download_workbook_data.click()
        download_button = self.driver.find_element_by_xpath(self.download_button_xpath)
        download_button.click()

# create a class to open quality overview session in stage from mysql
# incorporate this class and see how it works
# create another class for reporting mail
