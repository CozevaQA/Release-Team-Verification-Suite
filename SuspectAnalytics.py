from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, \
    ElementClickInterceptedException
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

# year and lob(optional) , no drill down



# dynamic xpath value for year selector,lob selector is not incorporated in class attributes

def create_connection(db_file):  # creating connection
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except Error as e:
        print(e)

    return None


class SuspectAnalytics:

    def __init__(self, driver):
        db_path = config.get("runner","dbpath")
        folder_path = ''
        self.conn = create_connection(db_path)
        self.cur = self.conn.cursor()
        self.driver = driver
        self.loader_element = 'sm_download_cssload_loader_wrap'
        self.selected_value_year_xpath = "//select[@name='year']//following::span[@class='multiselect-selected-text'][1]"
        self.service_year_xpath = "//select[@name='year']//following-sibling::div//child::button//child::span"
        self.drilldown_elements_xpath = "//div[@class=\"breadcrumb_dropdown\"]//child::a"
        self.apply_filter_xpath = "//a[text()=\"Apply\"]"
        self.select_all_id = 'sm_select_all'
        self.overview_xpath = "//div[@class=\"breadcrumb_dropdown\"]//child::a[1]"
        self.lob_xpath = "(//select[@name='lob_all']//following-sibling::div//child::button)[1]"
        self.lob_elements_xpath = "//select[@class='sm_dashboard_filter_select_single form-select' and @name='lob_all']//child::*"

    def action_click(self,element):
        try:
            element.click()
        except (ElementNotInteractableException,ElementClickInterceptedException):
            self.driver.execute_script("arguments[0].click();", element)

    def check_exists_byclass(self, classname):
        try:
            self.driver.find_element_by_class_name(classname)
        except NoSuchElementException:
            return False
        return True
    def browser_click(self,element):
        self.driver.execute_script("arguments[0].click();", element)


    def hasXpath(self,xpath):
        try:
            self.driver.find_element_by_xpath(xpath)
            return True
        except NoSuchElementException:
            return False

    def makedir(self, customer):
        path1 = str(customer) + "/" + str(customer) + "-SuspectAnalytics"
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
        WebDriverWait(self.driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
        runner.remove_chat_dashboard()
        if self.hasXpath(self.lob_xpath):
            for i in year:
                WebDriverWait(self.driver, 100).until(
                    EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                selected_value = self.driver.find_element_by_xpath(self.selected_value_year_xpath).text
                #print(selected_value)
                if int(selected_value) != int(i):
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
                                (int(customer), "Suspect Analytics", int(i), str("Does not exists"),
                                 str("Does not exists"),))
                        print(i,"does not exist ")
                        break
                print(i)
                lob = self.driver.find_element_by_xpath(self.lob_xpath)
                self.action_click(lob)
                lob_elements = self.driver.find_elements_by_xpath(self.lob_elements_xpath)
                print("Number of LOBs ", len(lob_elements))
                count = 1
                for lob_element in lob_elements:
                    WebDriverWait(self.driver, 100).until(
                        EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                    if (count > 1):
                        self.action_click(lob)
                    print(lob_element.get_attribute("value"))
                    val = lob_element.get_attribute("value")
                    lob_selector = "//label[@class=\"radio\"]//input[@value=\"%s\"]//following-sibling::span" % (val)
                    # print(st)
                    toclick=self.driver.find_element_by_xpath(lob_selector)
                    self.action_click(toclick)
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
                    # essential wait for loading page
                    WebDriverWait(self.driver, 100).until(
                        EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                    if self.check_exists_byclass("nodata"):
                        print("No data found in {} for lob {}".format(i, val))
                        try:
                            with self.conn:
                                self.cur.execute(
                                    'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                    (int(customer), "Suspect Analytics", int(i), str("Overview"),
                                     str(val)))
                                # print 'Data saved from exception'
                                ss_name = str(wbpath) + "/" + str(i) + str(
                                    val) + "Suspect Analytics" + "Overview" + ".png"  # naming convention is year-lob-drill.png
                                self.driver.save_screenshot(ss_name)
                        except sqlite3.IntegrityError:
                            pass
                    else:
                        ss_name = str(wbpath) + "/" + str(i) + str(
                            val) + "Suspect Analytics" + "Overview" + ".png"  # naming convention is year-lob-drill.png
                        self.driver.save_screenshot(ss_name)
                        print("data found")
        else:
            for i in year:
                # if self.check_exists_byclass("nodata"):
                #     print("No data found in {} ".format(i))
                #     try:
                #         with self.conn:
                #             self.cur.execute(
                #                 'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                #                 (int(customer), "Suspect Analytics", int(i), str("Overview"), "Medicare",))
                #             # print 'Data saved from exception'
                #             ss_name = str(wbpath) + "/" + str(
                #                 i) + "Suspect" + "Analytics" + ".png"  # naming convention is year-lob-drill.png
                #             self.driver.save_screenshot(ss_name)
                #     except sqlite3.IntegrityError:
                #         pass
                # else:
                #     ss_name = str(wbpath) + "/" + str(
                #         i) + "Suspect" + "Analytics" + ".png"  # naming convention is year-lob-drill.png
                #     self.driver.save_screenshot(ss_name)
                #     print("data found")
                WebDriverWait(self.driver, 200).until(
                    EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                selected_value = self.driver.find_element_by_xpath(self.selected_value_year_xpath).text
                #print("In without LOB ",selected_value)
                if int(selected_value) != int(i):
                    service_year = self.driver.find_element_by_xpath(self.service_year_xpath)
                    ActionChains(self.driver).move_to_element(service_year).click(service_year).perform()
                    year_selector = "//label[@class=\"radio\" and @title=\"%s\"]" % str(i)
                    try:
                        ele = self.driver.find_element_by_xpath(year_selector)
                        ele.location_once_scrolled_into_view
                        ele.click()
                    except NoSuchElementException:
                        with self.conn:
                            self.cur.execute(
                                'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                (int(customer), "Suspect Analytics", int(i), str("Does not exists"),
                                 str("Does not exists"),))
                        print(i, "does not exist ")
                        break
                print(i)

                if self.check_exists_byclass("nodata"):
                    print("No data found in {} ".format(i))
                    try:
                        with self.conn:
                            self.cur.execute(
                                'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                (int(customer), "Suspect Analytics", int(i), str("Overview"), "Medicare",))
                            # print 'Data saved from exception'
                            ss_name = str(wbpath) + "/" + str(
                                i) + "Suspect" + "Analytics" + ".png"  # naming convention is year-lob-drill.png
                            self.driver.save_screenshot(ss_name)
                    except sqlite3.IntegrityError:
                        pass
                else:
                    ss_name = str(wbpath) + "/" + str(
                        i) + "Suspect" + "Analytics" + ".png"  # naming convention is year-lob-drill.png
                    self.driver.save_screenshot(ss_name)
                    print("data found")
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





