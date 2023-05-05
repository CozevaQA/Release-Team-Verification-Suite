from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, \
    ElementClickInterceptedException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By

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
    loader=config.get("InpatientCost-Prod","loader_element")
    WebDriverWait(driver,100).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader)))


class InpatientCost:

    def __init__(self, driver):
        db_path = config.get("runner","dbpath")
        folder_path = ''
        self.conn = create_connection(db_path)
        self.cur = self.conn.cursor()
        self.driver = driver
        self.loader_element = config.get("InpatientCost-Prod","loader_element")
        self.selected_value_year_xpath = config.get("InpatientCost-Prod","selected_value_year_xpath")
        self.service_year_xpath = config.get("InpatientCost-Prod","service_year_xpath")
        self.lob_xpath = config.get("InpatientCost-Prod","lob_xpath")
        self.lob_elements_xpath = config.get("InpatientCost-Prod","lob_elements_xpath")

        self.apply_filter_xpath = config.get("InpatientCost-Prod","apply_filter_xpath")

        self.loader_panel_class = config.get("InpatientCost-Prod","loader_panel_class")
        self.drilldown_elements_xpath = config.get("InpatientCost-Prod","drilldown_elements_xpath")
        self.select_all_id = config.get("InpatientCost-Prod","select_all_id")
        self.overview_xpath = config.get("InpatientCost-Prod","overview_xpath")
        self.workbook_name= "Inpatient Cost"
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
        path1 = str(customer) + "/" + str(customer) + "-Inpatient Cost"
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
        for i in year:
            wait_to_load(self.driver)
            selected_value = self.driver.find_element_by_xpath(self.selected_value_year_xpath).text
            if int(selected_value) != i:

                service_year = self.driver.find_element_by_xpath(self.service_year_xpath)

                ActionChains(self.driver).move_to_element(service_year).click(service_year).perform()
                year_selector = "//label[@class=\"radio\" and @title=\"%s\"]" % str(i)
                try:
                    ele = self.driver.find_element_by_xpath(year_selector)
                    ele.location_once_scrolled_into_view
                    self.action_click(ele)
                except NoSuchElementException:
                    with self.conn:
                        self.cur.execute(
                            'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                            (int(customer), "Inpatient Cost", int(i), str("Does not exist"),
                             str("Does not  exists"),))
                        # print 'Data saved from exception'
                    print(i, "does not exist ")
                    break
            print(i)
            lob = self.driver.find_element_by_xpath(self.lob_xpath)
            self.action_click(lob)
            lob_elements = self.driver.find_elements_by_xpath(self.lob_elements_xpath)
            print("Number of LOBs ", len(lob_elements))
            count = 1
            for lob_element in lob_elements:
                wait_to_load(self.driver)
                if (count > 1):
                    self.action_click(lob)
                print(lob_element.get_attribute("value"))
                val = lob_element.get_attribute("value")
                lob_selector = "//label[@class=\"radio\"]//input[@value=\"%s\"]/following-sibling::span "% (val)
                # print(st)
                try:
                    self.action_click(self.driver.find_element_by_xpath(lob_selector))
                except ElementNotInteractableException:
                    b = self.driver.find_element_by_xpath(lob_selector)
                    self.driver.execute_script("arguments[0].click();", b)
                self.action_click(self.driver.find_element_by_xpath(self.apply_filter_xpath))
                # essential wait for loading page
                wait_to_load(self.driver)
                if self.check_exists_byclass("nodata"):
                    print("No data found in {} for lob {}".format(i, val))
                    try:
                        with self.conn:
                            self.cur.execute(
                                'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                (int(customer), "Inpatient Cost", int(i), str("By Service Category"), str(val),))
                            # print 'Data saved from exception'
                            ss_name = str(wbpath) + "/" + str(i) + str(
                                val) + "DX Category" + ".png"  # naming convention is year-lob-drill.png
                            self.driver.save_screenshot(ss_name)
                    except sqlite3.IntegrityError:
                        pass
                else:
                    ss_name = str(wbpath) + "/" + str(i) + str(
                        val) + "DX Category" + ".png"  # naming convention is year-lob-drill.png
                    self.driver.save_screenshot(ss_name)
                    drilldown_elements = self.driver.find_elements_by_xpath(self.drilldown_elements_xpath)
                    x = len(drilldown_elements)
                    b = 0
                    for j in range(2, x + 1):
                        # print("b value initially ,j  value initially ", b, j)
                        if b > 0:
                            continue
                        wait_to_load(self.driver)
                        self.action_click(self.driver.find_element_by_id(self.select_all_id))
                        drilldown_element = "//div[@class=\"breadcrumb_dropdown\"]//child::a[%s]" % (j)
                        # print(drilldown_element)
                        ele = self.driver.find_element_by_xpath(drilldown_element)
                        ele.click()
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
                                        (int(customer), str(self.workbook_name), int(i), str(drill_name), str(val),))
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
                                prev_drill.click()
                                wait_to_load(self.driver)
                                self.driver.find_element_by_id(self.select_all_id).click()
                                next_drill_xpath = "//div[@class=\"breadcrumb_dropdown\"]//child::a[%s]" % (b)
                                next_drill = self.driver.find_element_by_xpath(next_drill_xpath)
                                next_drill.click()
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
                                                    int(customer), str(self.workbook_name) ,int(i), str(drill_name2),
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
                    self.action_click(self.driver.find_element_by_xpath(self.overview_xpath))
                    count = count + 1


# create a class to open quality overview session in stage from mysql
# incorporate this class and see how it works
# create another class for reporting mail
