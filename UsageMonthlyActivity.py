from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, \
    ElementClickInterceptedException, UnexpectedAlertPresentException
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
    loader=config.get("Usage-MonthlyActivity-Prod","loader_element")
    try :
        WebDriverWait(driver, 1000).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader)))
    except UnexpectedAlertPresentException:
        print("Unknown Error Occurred while loading page ")





class UsageMonthlyActivity:

    def __init__(self, driver):
        db_path = db_path = config.get("runner","dbpath")
        folder_path = ''
        self.conn = create_connection(db_path)
        self.cur = self.conn.cursor()
        self.driver = driver
        self.loader_element = config.get("Usage-MonthlyActivity-Prod","loader_element")
        self.selected_value_year_xpath = config.get("Usage-MonthlyActivity-Prod","selected_value_year_xpath")
        self.service_year_xpath = config.get("Usage-MonthlyActivity-Prod","service_year_xpath")
        self.lob_xpath = config.get("Usage-MonthlyActivity-Prod","lob_xpath")
        self.apply_filter_xpath = config.get("Usage-MonthlyActivity-Prod","apply_filter_xpath")
        self.loader_panel_class = config.get("Usage-MonthlyActivity-Prod","loader_panel_class")
        self.drilldown_elements_xpath = config.get("Usage-MonthlyActivity-Prod", "drilldown_elements_xpath")
        self.lob_outer_elements_xpath = config.get("Usage-MonthlyActivity-Prod","lob_outer_elements_xpath")
        self.insert_chart_xpath=config.get("Usage-MonthlyActivity-Prod","insert_chart_xpath")
        self.select_all_id = config.get("Usage-MonthlyActivity-Prod", "select_all_id")
        self.overview_xpath = config.get("Usage-MonthlyActivity-Prod", "overview_xpath")
        self.activity_type_xpath=config.get("Usage-MonthlyActivity-Prod","activity_type_xpath")

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
        path1 = str(customer) + "/" + str(customer) + "-Usage Monthly Activity"
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
        lob_outer_elements = self.driver.find_elements_by_xpath(self.lob_outer_elements_xpath)
        count_outer = len(lob_outer_elements)
        #print(count_outer)
        loop_list = []
        count = 1
        loop_list.append(count_outer)
        lob_values = []
        for j in range(1, count_outer + 1):
            lob_inner_elements_xpath = self.lob_outer_elements_xpath + "[" + str(j) + "]//option"
            lob_inner_elements = self.driver.find_elements_by_xpath(lob_inner_elements_xpath)
            for k in range(1, len(lob_inner_elements) + 1):
                lob_inner_element_xpath = lob_inner_elements_xpath + "[" + str(k) + "]"
                #print(lob_inner_element_xpath)
                lob_value = self.driver.find_element_by_xpath(lob_inner_element_xpath).get_attribute("value")
                lob_values.append(lob_value)

            first_time = 1
            num_of_child = len(lob_inner_elements)
            loop_list.append(num_of_child)

        #print(loop_list)
        # print(lob_values)
        # 1. Click on Activity type
        # 2. From drop down extract list of elements
        # 3. For each element : select the element from the drop down and do the operations
        counter=1
        for counter in range(1,3):
            if (counter == 1):
                activity = "logon"
                print("Activity- Logon")
                logon_xpath = '//a[@title=" Logon"]//child::input'
                logon = self.driver.find_element_by_xpath(logon_xpath)
                self.action_click(logon)
            if(counter==2):

                activity_type = self.driver.find_element_by_xpath(self.activity_type_xpath)
                self.action_click(activity_type)
                activity = "Suppdata"
                print("Activity- Suppdata")
                suppdata_xpath = '//a[@title=" Supplemental Data"]//child::input'
                suppdata = self.driver.find_element_by_xpath(suppdata_xpath)
                self.action_click(suppdata)

            for y in year:
                wait_to_load(self.driver)
                selected_value = self.driver.find_element_by_xpath(self.selected_value_year_xpath).text
                if (customer == str('200')):
                    service_year_for_filter = y[0:4]
                else:
                    service_year_for_filter = y
                if int(selected_value) != service_year_for_filter:

                    service_year = self.driver.find_element_by_xpath(self.service_year_xpath)

                    ActionChains(self.driver).move_to_element(service_year).click(service_year).perform()
                    year_selector = "//label[@class=\"radio\" and @title=\"%s\"]" % str(service_year_for_filter)
                    try:
                        ele = self.driver.find_element_by_xpath(year_selector)
                        ele.location_once_scrolled_into_view
                        self.action_click(ele)
                    except NoSuchElementException:
                        with self.conn:
                            self.cur.execute(
                                'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                (int(customer), "Usage Monthly Activity", int(service_year_for_filter), str("Does not exist"),
                                 str("Does not  exists"),))
                            # print 'Data saved from exception'
                        print(service_year_for_filter, "does not exist ")
                        break
                print(y)

                j = 0
                count = 1
                for i in range(1, len(loop_list)):
                    count = 1
                    first_entry = 1
                    for x in range(1, loop_list[i] + 1):
                        if (count > 1):
                            lob = self.driver.find_element_by_xpath(self.lob_xpath)
                            self.driver.execute_script("arguments[0].click();", lob)
                        count = count + 1
                        if (first_entry == 1):
                            value_in_input = str(lob_values[j - 1])
                            val = value_in_input + activity
                            lobtobeclicked_xpath = "//input[@type=\"checkbox\" and @value=\"%s\"]//following-sibling::span" % (
                                value_in_input)
                            # print(lobtobeclicked_xpath)
                            lobtobeclicked = self.driver.find_element_by_xpath(lobtobeclicked_xpath)
                            self.action_click(lobtobeclicked)
                            first_entry = 0

                        value_in_input = str(lob_values[j])
                        print(value_in_input)
                        val = activity + value_in_input[0:4]
                        lobtobeclicked_xpath = "//input[@type=\"checkbox\" and @value=\"%s\"]//following-sibling::span" % (
                            value_in_input)
                        # print(lobtobeclicked_xpath)
                        lobtobeclicked = self.driver.find_element_by_xpath(lobtobeclicked_xpath)
                        self.action_click(lobtobeclicked)
                        j = j + 1
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
                        wait_to_load(self.driver)
                        if self.check_exists_byclass("nodata"):
                            print("No data found in {} for lob {}".format(y, val))
                            try:
                                with self.conn:
                                    self.cur.execute(
                                        'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',
                                        (int(customer), "Usage-Monthly Activity", int(service_year_for_filter), str("Overview"),
                                         str(val),))
                                    # print 'Data saved from exception'
                                    ss_name = str(wbpath) + "/" + str(service_year_for_filter) + str(
                                        activity+value_in_input[0:4]) + ".png"  # naming convention is year-lob-drill.png
                                    print(ss_name)
                                    self.driver.save_screenshot(ss_name)
                            except sqlite3.IntegrityError:
                                pass


                        else:
                            ss_name = str(wbpath) + "/" + str(service_year_for_filter) + str(
                                activity + value_in_input[0:4]) + ".png"  # naming convention is year-lob-drill.png
                            print(ss_name)
                            self.driver.save_screenshot(ss_name)
                            self.driver.find_element_by_xpath(self.insert_chart_xpath).click()
                            drilldown_elements = self.driver.find_elements_by_xpath(self.drilldown_elements_xpath)

                            x = len(drilldown_elements)

                            b = 0

                            for r in range(2, x + 1):

                                # print("b value initially ,j  value initially ", b, j)

                                if b > 0:
                                    continue

                                wait_to_load(self.driver)

                                self.driver.find_element_by_id(self.select_all_id).click()

                                drilldown_element = "//div[@class=\"breadcrumb_dropdown\"]//child::a[%s]" % (r)

                                # print(drilldown_element)

                                ele = self.driver.find_element_by_xpath(drilldown_element)

                                self.action_click(ele)

                                drill_name = ele.get_attribute("drill_down_name")

                                print(drill_name)

                                wait_to_load(self.driver)

                                if self.check_exists_byclass("nodata"):

                                    # print("No data found first , value of j is ", j)

                                    print(

                                        "No data found in year {} , drill down {} for lob {} ".format(service_year_for_filter, drill_name,
                                                                                                      val))

                                    try:

                                        with self.conn:

                                            self.cur.execute(

                                                'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',

                                                (int(customer), "Usage-Monthly Activity", str(service_year_for_filter), str(drill_name),
                                                 str(val),))

                                            # print 'Data saved from exception'

                                            ss_name = str(wbpath) + "/" + str(service_year_for_filter) + str(val) + str(

                                                drill_name) + ".png"  # naming convention is year-lob-drill.png

                                            self.driver.save_screenshot(ss_name)

                                    except sqlite3.IntegrityError:

                                        pass

                                    a = r - 1

                                    b = r + 1

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

                                                "No data found in year {} , drill down {} for lob {} ".format(str(service_year_for_filter),
                                                                                                              drill_name2,

                                                                                                              val))

                                            try:

                                                with self.conn:

                                                    self.cur.execute(

                                                        'INSERT INTO analytics_nodata_found (Customer,Workbook,Year,DrillDown,LOB) VALUES (?,?,?,?,?)',

                                                        (

                                                            int(customer), "Usage-Monthly Activity", str(service_year_for_filter),
                                                            str(drill_name2),

                                                            str(val),))

                                                    # print 'Data saved from exception'

                                                    ss_name = str(wbpath) + "/" + str(service_year_for_filter) + str(val) + str(

                                                        drill_name2) + ".png"  # naming convention is year-lob-drill.png

                                                    self.driver.save_screenshot(ss_name)

                                            except sqlite3.IntegrityError:

                                                pass

                                        else:
                                            ss_name = str(wbpath) + "/" + str(service_year_for_filter) + str(val) + str(

                                                drill_name2) + ".png"  # naming convention is year-lob-drill.png

                                            self.driver.save_screenshot(ss_name)

                                            print("data found")

                                        b = b + 1

                                        # print("checking increment ! b and j is {} , {} respectively", b, j)


                                else:

                                    ss_name = str(wbpath) + "/" + str(service_year_for_filter) + str(val) + str(

                                        drill_name) + ".png"  # naming convention is year-lob-drill.png

                                    self.driver.save_screenshot(ss_name)
                                    print("data found")

                                    r = r + 1

                                # print("j value in the end ",j)

                            loader_element = 'sm_download_cssload_loader_wrap'

                            wait_to_load(self.driver)

                            self.driver.find_element_by_xpath(self.overview_xpath).click()

                            count = count + 1

        counter=counter+1













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










# create a class to open Usage-Monthly Activity session in stage from mysql
# incorporate this class and see how it works
# create another class for reporting mail
