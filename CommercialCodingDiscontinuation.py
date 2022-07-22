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

#year , lob(optional) and drill down

# dynamic xpath value for year selector,lob selector is not incorporated in class attributes

def create_connection(db_file):  # creating connection
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except Error as e:
        print(e)

    return None


class CommercialCodingDiscontinuation:

    def __init__(self, driver):
        db_path = config.get("runner","dbpath")
        folder_path = ''
        self.conn = create_connection(db_path)
        self.cur = self.conn.cursor()
        self.driver = driver
        self.loader_element = 'sm_download_cssload_loader_wrap'
        self.selected_value_year_xpath = '//select[@name="year"]//following::span[@class="multiselect-selected-text"][1]'
        self.service_year_xpath = '//select[@name="year"]//following-sibling::div//child::button//child::span'
        self.drilldown_elements_xpath = "//div[@class=\"breadcrumb_dropdown\"]//child::a"
        self.apply_filter_xpath = "//a[text()=\"Apply\"]"
        self.select_all_id = 'sm_select_all'
        self.overview_xpath = "//div[@class=\"breadcrumb_dropdown\"]//child::a[1]"
        self.lob_xpath = '(//select[@name="lob_all"]//following-sibling::div//child::button)[1]'
        self.lob_select_all_xpath = '//select[ @name="lob_all"]//following::input[@value="multiselect-all"][1]//following-sibling::span[contains(text()," Select all")]'
        self.lob_elements_xpath = '//select[@name="lob_all"]//child::*'


    def check_exists_byclass(self, classname):
        try:
            self.driver.find_element_by_class_name(classname)
        except NoSuchElementException:
            return False
        return True

    def makedir(self, customer):
        path1 = str(customer) + "/" + str(customer) + "-CommercialCodingDiscontinuation"
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
    def action_click(self,element):
        try:
            element.click()
        except (ElementNotInteractableException,ElementClickInterceptedException):
            self.driver.execute_script("arguments[0].click();", element)

    def hasXpath(self,xpath):
        try:
            self.driver.find_element_by_xpath(xpath)
            return True
        except NoSuchElementException:
            return False


    def iterate_filter(self, year, customer):
        wbpath = self.makedir(customer)
        WebDriverWait(self.driver, 100).until(EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
        if self.hasXpath(self.lob_xpath):
            for i in year:
                WebDriverWait(self.driver, 100).until(
                    EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
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
                                (int(customer), "Commercial CodingDiscontinuation", int(i), str("Does not exist"), str("Does not  exists"),))
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
                    WebDriverWait(self.driver, 100).until(
                        EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                    sel_all = self.driver.find_element_by_xpath(self.lob_select_all_xpath)
                    if (count > 1):
                        self.action_click(lob)
                    if (len(lob_elements) == 1):
                        self.action_click(sel_all)
                    else:
                        self.action_click(sel_all)
                        self.action_click(sel_all)
                    print(lob_element.get_attribute("value"))
                    val = lob_element.get_attribute("value")
                    lob_selector = "//label[@class=\"checkbox\"]//input[@value=\"%s\"]/following-sibling::span" % (val)
                    # print(st)
                    self.action_click(self.driver.find_element_by_xpath(lob_selector))
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
                    self.driver.find_element_by_xpath(self.apply_filter_xpath).click()
                    # essential wait for loading page
                    WebDriverWait(self.driver, 100).until(
                        EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                    if self.check_exists_byclass("nodata"):
                        print("No data found in {} for lob {}".format(i, val))
                        try:
                            with self.conn:
                                self.cur.execute(
                                    'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                    (int(customer), "Commercial CodingDiscontinuation", int(i), str("Overview"), str(val),))
                                # print 'Data saved from exception'
                                ss_name = str(wbpath) + "/" + str(i) + str(
                                    val) +"Commercial-Risk CodingDiscontinuation"+ "Overview" + ".png"  # naming convention is year-lob-drill.png
                                self.driver.save_screenshot(ss_name)
                        except sqlite3.IntegrityError:
                            pass
                    else:
                        print("data found")
                        ss_name = str(wbpath) + "/" + str(i) + str(
                            val) + "Commercial-Risk CodingDiscontinuation" + "Overview" + ".png"  # naming convention is year-lob-drill.png
                        self.driver.save_screenshot(ss_name)
                        drilldown_elements = self.driver.find_elements_by_xpath(self.drilldown_elements_xpath)
                        x = len(drilldown_elements)
                        b = 0
                        for j in range(2, x + 1):
                            # print("b value initially ,j  value initially ", b, j)
                            if b > 0:
                                continue
                            WebDriverWait(self.driver, 100).until(
                                EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                            self.driver.find_element_by_id(self.select_all_id).click()
                            drilldown_element = "//div[@class=\"breadcrumb_dropdown\"]//child::a[%s]" % (j)
                            # print(drilldown_element)
                            ele = self.driver.find_element_by_xpath(drilldown_element)
                            ele.click()
                            drill_name = ele.get_attribute("drill_down_name")
                            print(drill_name)
                            WebDriverWait(self.driver, 100).until(
                                EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                            if self.check_exists_byclass("nodata"):
                                # print("No data found first , value of j is ", j)
                                print(
                                    "No data found in year {} , drill down {} for lob {} ".format(i, drill_name, val))
                                try:
                                    with self.conn:
                                        self.cur.execute(
                                            'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                            (int(customer), "Commercial-CodingDiscontinuation", int(i), str(drill_name), str(val),))
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
                                    WebDriverWait(self.driver, 100).until(
                                        EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                                    self.driver.find_element_by_id(self.select_all_id).click()
                                    next_drill_xpath = "//div[@class=\"breadcrumb_dropdown\"]//child::a[%s]" % (b)
                                    next_drill = self.driver.find_element_by_xpath(next_drill_xpath)
                                    next_drill.click()
                                    drill_name2 = next_drill.get_attribute("drill_down_name")
                                    print(drill_name2)
                                    WebDriverWait(self.driver, 100).until(
                                        EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                                    if self.check_exists_byclass("nodata"):
                                        # print("No data found second , value of b is {} and j is {} is ", format(b, j))
                                        print(
                                            "No data found in year {} , drill down {} for lob {} ".format(i,
                                                                                                          drill_name2,
                                                                                                          val))
                                        try:
                                            with self.conn:
                                                self.cur.execute(
                                                    'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                                    (
                                                        int(customer), "Commercial-CodingDiscontinuation", int(i), str(drill_name2),
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
                        WebDriverWait(self.driver, 100).until(
                            EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
                        self.driver.find_element_by_xpath(self.overview_xpath).click()
                        count = count + 1


        else:
            for i in year:
                WebDriverWait(self.driver, 100).until(
                    EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                selected_value = self.driver.find_element_by_xpath(self.selected_value_year_xpath).text
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
                                (int(customer), "Commercial CodingDiscontinuation", int(i), str("Does not exist"), str("Does not  exists"),))
                            # print 'Data saved from exception'
                        print(i, "does not exist ")
                        break
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
                    self.driver.find_element_by_xpath(self.apply_filter_xpath).click()

                else:
                    print(i)
                    if self.check_exists_byclass("nodata"):
                        print("No data found in {} ".format(i))
                        try:
                            with self.conn:
                                self.cur.execute(
                                    'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                    (int(customer), "Commercial CodingDiscontinuation", int(i), str("Overview"),
                                     "Commercial",))
                                # print 'Data saved from exception'
                                ss_name = str(wbpath) + "/" + str(
                                    i) + "Commercial" + "CodingDiscontinuation" + ".png"  # naming convention is year-lob-drill.png
                                self.driver.save_screenshot(ss_name)
                        except sqlite3.IntegrityError:
                            pass

                    else:
                        drilldown_elements = self.driver.find_elements_by_xpath(self.drilldown_elements_xpath)
                        x = len(drilldown_elements)
                        b = 0
                        for j in range(2, x + 1):
                            # print("b value initially ,j  value initially ", b, j)
                            if b > 0:
                                continue
                            WebDriverWait(self.driver, 100).until(
                                EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                            self.driver.find_element_by_id(self.select_all_id).click()
                            drilldown_element = "//div[@class=\"breadcrumb_dropdown\"]//child::a[%s]" % (j)
                            # print(drilldown_element)
                            ele = self.driver.find_element_by_xpath(drilldown_element)
                            ele.click()
                            drill_name = ele.get_attribute("drill_down_name")
                            print(drill_name)
                            WebDriverWait(self.driver, 100).until(
                                EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                            if self.check_exists_byclass("nodata"):
                                # print("No data found first , value of j is ", j)
                                print(
                                    "No data found in year {} , drill down {} for lob {} ".format(i, drill_name,
                                                                                                  "Commercial"))
                                try:
                                    with self.conn:
                                        self.cur.execute(
                                            'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                            (int(customer), "Commercial RiskOverview", int(i), str(drill_name),
                                             "Commercial",))
                                        # print 'Data saved from exception'
                                        ss_name = str(wbpath) + "/" + str(i) + "Commercial" + str(
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
                                    WebDriverWait(self.driver, 100).until(
                                        EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                                    self.driver.find_element_by_id(self.select_all_id).click()
                                    next_drill_xpath = "//div[@class=\"breadcrumb_dropdown\"]//child::a[%s]" % (b)
                                    next_drill = self.driver.find_element_by_xpath(next_drill_xpath)
                                    next_drill.click()
                                    drill_name2 = next_drill.get_attribute("drill_down_name")
                                    print(drill_name2)
                                    WebDriverWait(self.driver, 100).until(
                                        EC.invisibility_of_element_located((By.CLASS_NAME, self.loader_element)))
                                    if self.check_exists_byclass("nodata"):
                                        # print("No data found second , value of b is {} and j is {} is ", format(b, j))
                                        print(
                                            "No data found in year {} , drill down {} for lob {} ".format(i,
                                                                                                          drill_name2,
                                                                                                          "Commercial"))
                                        try:
                                            with self.conn:
                                                self.cur.execute(
                                                    'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                                    (
                                                        int(customer), "Quality Overview", int(i), str(drill_name2),
                                                        "Commercial",))
                                                # print 'Data saved from exception'
                                                ss_name = str(wbpath) + "/" + str(i) + "Commercial" + str(
                                                    drill_name2) + ".png"  # naming convention is year-lob-drill.png
                                                self.driver.save_screenshot(ss_name)
                                        except sqlite3.IntegrityError:
                                            pass
                                    else:
                                        ss_name = str(wbpath) + "/" + str(i) + "Commercial" + str(
                                            drill_name2) + ".png"  # naming convention is year-lob-drill.png
                                        self.driver.save_screenshot(ss_name)
                                        print("data found")

                                    b = b + 1
                                    # print("checking increment ! b and j is {} , {} respectively", b, j)

                            else:
                                ss_name = str(wbpath) + "/" + str(i) + "Commercial" + str(
                                    drill_name) + ".png"  # naming convention is year-lob-drill.png
                                self.driver.save_screenshot(ss_name)
                                print("data found")
                                j = j + 1
                            # print("j value in the end ",j)
                        loader_element = 'sm_download_cssload_loader_wrap'
                        WebDriverWait(self.driver, 100).until(
                            EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
                        self.driver.find_element_by_xpath(self.overview_xpath).click()


























