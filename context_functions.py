import re
import traceback
from random import randint

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, \
    ElementClickInterceptedException
from sigfig import round

import support_functions as sf
import variablestorage as locator
import time

global global_search_prov, global_search_prac, global_search_pat


def init_global_search():
    global global_search_pat
    global global_search_prov
    global global_search_prac
    global_search_prov = None
    global_search_pat = None
    global_search_prac = None


def support_menubar(driver, workbook, ws, logger, run_from):
    if ws is None:
        workbook.create_sheet('Support Menubar')
        ws = workbook['Support Menubar']

    ws.append(['ID', 'Context', 'Scenario', 'Status', 'Time Taken', 'Comments'])
    logger.info("MenubarNavigation function started.")
    header_font = Font(color='FFFFFF', bold=False, size=12)
    header_cell_color = PatternFill('solid', fgColor='030303')
    ws['A1'].font = header_font
    ws['A1'].fill = header_cell_color
    ws['B1'].font = header_font
    ws['B1'].fill = header_cell_color
    ws['C1'].font = header_font
    ws['C1'].fill = header_cell_color
    ws['D1'].font = header_font
    ws['D1'].fill = header_cell_color
    ws['E1'].font = header_font
    ws['E1'].fill = header_cell_color
    ws.name = "Arial"
    test_case_id = 1
    main_registry_url = driver.current_url

    try:
        logger.info("Menubar navigation block started.")
        time.sleep(1)
        sf.ajax_preloader_wait(driver)
        current_url = driver.current_url
        access_message = sf.CheckAccessDenied(current_url)
        error_message = sf.CheckErrorMessage(driver)

        if access_message == 0 and error_message == 0:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_side_nav_SlideOut)))
            time.sleep(0.5)
            context_name = driver.find_element_by_xpath(locator.xpath_context_Name).text
            print(context_name)
            test_case_id += 1
            ws.append((test_case_id, '', 'Set context to: ' + context_name, 'Passed'))
            logger.info("if#1 block ended.")

        elif access_message == 1:
            logger.info("elif#1 block started.")
            test_case_id += 1
            ws.append([test_case_id, '', 'Access Check', 'Failed'])
            logger.info("elif#1 block ended.")

        elif error_message == 1:
            logger.info("elif#2 block started.")
            test_case_id += 1
            ws.append((test_case_id, '', 'Default context without error message', 'Failed'))
            logger.info("elif#2 block ended.")

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, locator.xpath_side_nav_SlideOut)))
        time.sleep(0.5)
        context_name = driver.find_element_by_xpath(locator.xpath_context_Name).text
        print(context_name)
        driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()

        try:
            time.sleep(1.5)
            links = driver.find_elements_by_xpath(locator.xpath_menubar_Item_Link)
            names = driver.find_elements_by_xpath(locator.xpath_menubar_Item_Link_Name)
            driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()

        except Exception as e:
            test_case_id += 1
            ws.append((test_case_id, context_name, 'Sidemenubar navigation', 'Failed'))
            print(e)
            traceback.print_exc()

        for link in range(len(links)) and range(len(names)):
            # time.sleep(1.5)
            driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()
            time.sleep(0.5)
            driver.execute_script("arguments[0].scrollIntoView();", links[link])
            print("Link Index: " + str(link))
            print(names[link].text)
            link_name = names[link].text
            try:
                links[link].click()
                start_time = time.perf_counter()
                sf.ajax_preloader_wait(driver)
                total_time = time.perf_counter() - start_time
                current_url = driver.current_url
                access_message = sf.CheckAccessDenied(current_url)
                if access_message == 1:
                    print("Access Denied found!")
                    logger.error(context_name + "-->" + link_name + ": " + "Access Denied found!")
                    test_case_id += 1
                    ws.append((test_case_id, context_name, 'Access Check for ' + link_name, 'Failed'))
                else:
                    print("Access Check done!")
                    error_message = sf.CheckErrorMessage(driver)
                    if error_message == 1:
                        print("Error toast message is displayed")
                        # logger.critical("ERROR TOAST MESSAGE IS DISPLAYED!")
                        test_case_id += 1
                        ws.append((test_case_id, context_name, link_name + ' without error message', 'Failed'))
                        logger.error(context_name + "-->" + link_name + ": " + "Error message found!")
                    else:
                        if len(driver.find_elements_by_xpath(locator.xpath_data_Table_Info)) != 0:
                            time.sleep(0.5)
                            datatable_info = driver.find_element_by_xpath(locator.xpath_data_Table_Info).text
                            print(datatable_info)
                            test_case_id += 1
                            if link_name == "Providers":
                                x=1

                            ws.append((test_case_id, context_name, 'Navigation to ' + link_name, 'Passed',
                                       str(round(total_time, sigfigs=3)),
                                       datatable_info))
                            logger.info(context_name + "-->" + link_name + ": " + "Navigation done.")

                        else:
                            print("No datatable!")
                            test_case_id += 1
                            ws.append((test_case_id, context_name, 'Navigation to ' + link_name, 'Passed',
                                       str(round(total_time, sigfigs=3))))

            except Exception as e:
                print(e)
                traceback.print_exc()
                test_case_id += 1
                ws.append((test_case_id, context_name, 'Navigation to ' + link_name, 'Failed'))

            finally:
                links = driver.find_elements_by_xpath(locator.xpath_menubar_Item_Link)
                names = driver.find_elements_by_xpath(locator.xpath_menubar_Item_Link_Name)

        driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, locator.xpath_registry_Link)))
        time.sleep(2)
        driver.find_element_by_xpath(locator.xpath_registry_Link).click()
        sf.ajax_preloader_wait(driver)

    except Exception as e:
        print(e)
        traceback.print_exc()
        test_case_id += 1
        ws.append((test_case_id, "", 'Menubar Navigation', 'Failed'))

    rows = ws.max_row
    cols = ws.max_column
    for i in range(2, rows + 1):
        for j in range(3, cols + 1):
            if ws.cell(i, j).value == 'Passed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='0FC404')
            elif ws.cell(i, j).value == 'Failed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FC0E03')
            elif ws.cell(i, j).value == 'Showing 0 to 0':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FCC0BB')


def practice_menubar(driver, workbook, logger, run_from):
    workbook.create_sheet('Practice Menubar')
    ws = workbook['Practice Menubar']
    main_registry_url = driver.current_url
    if run_from == "Cozeva Support" or run_from == "Limited Cozeva Support" or run_from == "Customer Support" or run_from == "Regional Support":
        # Switching to random Practice name from default set context, main page
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_side_nav_SlideOut)))
            driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()
            driver.find_element_by_id("providers-list").click()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'tabs')))
            driver.find_element_by_class_name("tabs").find_elements_by_tag_name('li')[0].click()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "metric-support-prac-ls")))
            list_of_practice_elements = driver.find_element_by_id("metric-support-prac-ls").find_elements_by_tag_name(
                'tr')
            selected_practice = list_of_practice_elements[
                sf.RandomNumberGenerator(len(list_of_practice_elements), 1)[0]].find_element_by_tag_name('a')
            global global_search_prac
            global_search_prac = selected_practice.text
            selected_practice.click()
        except Exception as e:
            ws.append(['1', "Attempting to navigate to a random practice", 'Navigation to practice context', 'Failed',
                       "Unable to navigate to a practice. Either the Practice list is unreachable or navigation access is denied"])
            print(e)
            traceback.print_exc()
            return
    elif run_from == "Office Admin Provider Delegate" or run_from == "Provider":
        ws.append(["1", run_from + " Role does not have access to practice Submenus"])
        return
    support_menubar(driver, workbook, ws, logger, run_from)

    driver.get(main_registry_url)
    sf.ajax_preloader_wait(driver)


def provider_menubar(driver, workbook, logger, run_from):
    workbook.create_sheet('Provider Menubar')
    ws = workbook['Provider Menubar']
    main_registry_url = driver.current_url
    if run_from == "Cozeva Support" or run_from == "Limited Cozeva Support" or run_from == "Customer Support" or run_from == "Regional Support" or run_from == "Office Admin Practice Delegate":
        # Switching to random Provider name from default set context, main page
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_side_nav_SlideOut)))
            driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()
            driver.find_element_by_id("providers-list").click()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.ID, 'metric-support-prov-ls')))
            list_of_provider_elements = driver.find_element_by_id("metric-support-prov-ls").find_elements_by_tag_name(
                'tr')
            selected_provider = list_of_provider_elements[
                sf.RandomNumberGenerator(len(list_of_provider_elements), 1)[0]].find_elements_by_tag_name('a')[1]
            global global_search_prov
            global_search_prov = selected_provider.text
            selected_provider.click()
        except Exception as e:
            ws.append(['1', "Attempting to navigate to a random provider", 'Navigation to provider context', 'Failed',
                       "Unable to navigate to a provider. Either the Provider list is unreachable or navigation access is denied"])
            driver.get(main_registry_url)
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_side_nav_SlideOut)))
            print(e)
            traceback.print_exc()
            return

    support_menubar(driver, workbook, ws, logger, run_from)

    # if run_from == "Cozeva Support" or run_from == "Limited Cozeva Support" or run_from == "Customer Support" or run_from == "Regional Support":
    #     driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()
    #     driver.find_element_by_id("home").click()
    #     sf.ajax_preloader_wait(driver)
    driver.get(main_registry_url)
    sf.ajax_preloader_wait(driver)
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, locator.xpath_side_nav_SlideOut)))


def patient_dashboard(driver, workbook, logger, run_from):
    workbook.create_sheet('Patient Dashboard')
    ws = workbook['Patient Dashboard']
    ws.append(['ID', 'Context', 'Scenario', 'Status', 'Time Taken', 'Comments'])
    header_font = Font(color='FFFFFF', bold=False, size=12)
    header_cell_color = PatternFill('solid', fgColor='030303')
    ws['A1'].font = header_font
    ws['A1'].fill = header_cell_color
    ws['B1'].font = header_font
    ws['B1'].fill = header_cell_color
    ws['C1'].font = header_font
    ws['C1'].fill = header_cell_color
    ws['D1'].font = header_font
    ws['D1'].fill = header_cell_color
    ws.name = "Arial"
    test_case_id = 1

    def hoverCheck(driver, ws, run_from, Pcp_hover, test_case_id):
        x = 1

    # From Starting point Registry, navigate to a random patient of a random metric
    main_registry_url = driver.current_url
    window_switched = 0
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "registry_body")))
        metrics = driver.find_element_by_id("registry_body").find_elements_by_tag_name('li')
        print("Provider Registry metrics loaded into a variable")
        percent = '0.00'
        while percent == '0.00' or percent == '0.00%':
            if len(metrics) > 1:
                selectedMetric = metrics[sf.RandomNumberGenerator(len(metrics), 1)[0]]
                percent = selectedMetric.find_element_by_class_name('percent').text
            else:
                selectedMetric = metrics[0]
                percent = selectedMetric.find_element_by_class_name('percent').text
        print("Found a Suitable Metric to click on")
        print("Attempting to click on " + selectedMetric.text)
        selectedMetric.click()
        print("Click Performed")
        sf.ajax_preloader_wait(driver)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'tabs')))

        if run_from == "Cozeva Support" or run_from == "Customer Support" or run_from == "Regional Support" or run_from == "Limited Cozeva Support":
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'tabs')))
            driver.find_element_by_class_name('tabs').find_elements_by_class_name('tab')[2].click()
            sf.ajax_preloader_wait(driver)
            if len(driver.find_elements_by_class_name('dt_tag_value')) > 0:
                driver.find_element_by_class_name('dt_tag_close').click()
                sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "metric-support-pat-ls")))
            patients = driver.find_element_by_id("metric-support-pat-ls").find_elements_by_tag_name('tr')
            patients[sf.RandomNumberGenerator(len(patients), 1)[0]].find_element_by_class_name('pat_name').click()
        elif run_from == "Office Admin Practice Delegate":
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'tabs')))
            driver.find_element_by_class_name('tabs').find_elements_by_class_name('tab')[1].click()
            sf.ajax_preloader_wait(driver)
            if len(driver.find_elements_by_class_name('dt_tag_value')) > 0:
                driver.find_element_by_class_name('dt_tag_close').click()
                sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "metric-support-pat-ls")))
            patients = driver.find_element_by_id("metric-support-pat-ls").find_elements_by_tag_name('tr')
            patients[sf.RandomNumberGenerator(len(patients), 1)[0]].find_element_by_class_name('pat_name').click()
        elif run_from == "Office Admin Provider Delegate" or run_from == "Provider":
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'tabs')))
            if len(driver.find_elements_by_class_name('dt_tag_value')) > 0:
                driver.find_element_by_class_name('dt_tag_close').click()
                sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "quality_registry_list")))
            patients = driver.find_element_by_id("quality_registry_list").find_elements_by_tag_name('tr')
            patients[sf.RandomNumberGenerator(len(patients), 1)[0]].find_element_by_class_name('pat_name').click()

        driver.switch_to.window(driver.window_handles[1])
        window_switched = 1
        sf.ajax_preloader_wait(driver)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, locator.xpath_cozeva_Id)))
        patient_id = driver.find_element_by_xpath(locator.xpath_cozeva_Id).text
        global global_search_pat
        global_search_pat = patient_id
        current_url = driver.current_url
        access_message = sf.CheckAccessDenied(current_url)

        if access_message == 1:
            print("Access Denied found!")
            # logger.critical("Access Denied found!")
            test_case_id += 1
            ws.append((test_case_id, patient_id, 'Navigation to dashboard page',
                       'Failed', 'x', 'Access Denied'))

        else:
            print("Access Check done!")
            # logger.info("Access Check done!")
            error_message = sf.CheckErrorMessage(driver)

            if error_message == 1:
                print("Error toast message is displayed")
                # logger.critical("ERROR TOAST MESSAGE IS DISPLAYED!")
                test_case_id += 1
                ws.append \
                    ((test_case_id, patient_id, 'Navigation to dashboard page ',
                      'Failed', 'x', 'Error toast message is displayed'))

            else:
                measure_count_dashboard = len \
                    (driver.find_elements_by_xpath("//tbody[@class='measurement-body careops-new']/tr"))
                test_case_id += 1
                ws.append((test_case_id, patient_id, 'Navigation to dashboard page',
                           'Passed', 'x', 'Measures count in dashboard: ' + str(measure_count_dashboard)))
                logger.info(patient_id + ": Navigated to patient dashboard.")
                """ **** PCP INFO BLOCK **** """
                try:
                    Pcp_Name = driver.find_element_by_id("pcp_name").text
                    Pcp_Webelement = driver.find_element_by_id("pcp_name")
                    Pcp_hover = Pcp_Webelement.get_attribute("data-tooltip")
                    # Pcp_Name = "-"
                    # Pcp_hover = "N/A, N/A, N/A"

                    if Pcp_Name == '-':
                        test_case_id += 1
                        ws.append((test_case_id, patient_id, 'PCP Name',
                                   'Failed', 'x', "PCP Name is Blank"))
                    elif Pcp_Name == "N/A":
                        test_case_id += 1
                        ws.append((test_case_id, patient_id, 'PCP Name',
                                   'Failed', 'x', "PCP Name is NA"))
                    else:
                        test_case_id += 1
                        ws.append((test_case_id, patient_id, 'PCP Name',
                                   'Passed', 'x', Pcp_Name))

                    if Pcp_hover == "N/A, N/A, No Practice":
                        test_case_id += 1
                        ws.append((test_case_id, patient_id, 'PCP Attribution on hover',
                                   'Failed', 'x', "PCP does not have Region/Panel Attribution"))
                    elif Pcp_hover == "N/A, N/A, N/A":
                        test_case_id += 1
                        ws.append((test_case_id, patient_id, 'PCP Attribution on hover',
                                   'Failed', 'x', "PCP does not have any attribution"))
                    else:
                        test_case_id += 1
                        ws.append((test_case_id, patient_id, 'PCP Attribution on hover',
                                   'Passed', 'x', Pcp_hover))
                        # function for hovercheck
                        hoverCheck(driver, ws, run_from, Pcp_hover, test_case_id)
                except Exception as e:
                    print(e)
                    traceback.print_exc()
                    test_case_id += 1
                    ws.append((test_case_id, patient_id, 'PCP hover',
                               'Failed', 'x', "PCP Name is not present/Not interactable"))

                # Aspy Edit ------------------------------------------------------------------------------------
                # '''
                """ **** PATIENT MENUBAR NAVIGATION **** """
                WebDriverWait(driver, 30).until(EC.presence_of_element_located(
                    (By.XPATH, locator.xpath_patient_Header_Dropdown_Arrow)))
                driver.find_element_by_xpath(locator.xpath_patient_Header_Dropdown_Arrow).click()
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, locator.xpath_patient_History)))
                history_link = driver.find_element_by_xpath(locator.xpath_patient_History)
                ActionChains(driver).move_to_element(history_link).perform()
                patient_history_items = driver.find_elements_by_xpath(
                    locator.xpath_patient_History_Item_Link)
                driver.find_element_by_xpath(locator.xpath_patient_Header_Dropdown_Arrow).click()
                for item_counter in range(len(patient_history_items)):
                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located(
                            (By.XPATH, locator.xpath_patient_Header_Dropdown_Arrow)))
                    driver.find_element_by_xpath(locator.xpath_patient_Header_Dropdown_Arrow).click()
                    WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((By.XPATH, locator.xpath_patient_History)))
                    history_link = driver.find_element_by_xpath(locator.xpath_patient_History)
                    time.sleep(0.5)
                    ActionChains(driver).move_to_element(history_link).perform()
                    item_name = patient_history_items[item_counter].text
                    print(item_name)
                    time.sleep(0.5)
                    patient_history_items[item_counter].click()
                    start_time = time.perf_counter()
                    sf.ajax_preloader_wait(driver)
                    total_time = time.perf_counter() - start_time
                    current_url = driver.current_url
                    access_message = sf.CheckAccessDenied(current_url)

                    if access_message == 1:
                        print("Access Denied found!")
                        # logger.critical("Access Denied found!")
                        test_case_id += 1
                        ws.append((test_case_id,
                                   patient_id, 'Navigation to  ' + item_name,
                                   'Failed', 'x', 'Access Denied'))

                    else:
                        print("Access Check done!")
                        # logger.info("Access Check done!")
                        error_message = sf.CheckErrorMessage(driver)

                        if error_message == 1:
                            print("Error toast message is displayed")
                            # logger.critical("ERROR TOAST MESSAGE IS DISPLAYED!")
                            test_case_id += 1
                            ws.append((test_case_id,
                                       patient_id, 'Navigation to  ' + item_name,
                                       'Failed', 'x', 'Error toast message is displayed'))

                        else:

                            if len(driver.find_elements_by_xpath(locator.xpath_data_Table_Row)) != 0:
                                if len(driver.find_elements_by_xpath(
                                        locator.xpath_empty_Data_Table_Row)) != 0:
                                    print("Data table is empty")
                                    test_case_id += 1
                                    ws.append((test_case_id, patient_id, "Navigation to" + item_name, 'Passed',
                                               round(total_time, sigfigs=3),
                                               'Data table is empty'))
                                    logger.info("Navigated to:  " + item_name)
                                else:
                                    table_row_count = len(
                                        driver.find_elements_by_xpath(locator.xpath_data_Table_Row))
                                    print(table_row_count)
                                    test_case_id += 1
                                    ws.append((test_case_id, patient_id, "Navigation to" + item_name, 'Passed',
                                               round(total_time, sigfigs=3),
                                               'Data table row count in the first page: ' + str
                                               (table_row_count)))
                                    logger.info("Navigated to: " + item_name)

                            patient_history_items = driver.find_elements_by_xpath(
                                locator.xpath_patient_History_Item_Link)

                """ **** PATIENT INFO **** """
                driver.find_element_by_xpath(locator.xpath_patient_Header_Dropdown_Arrow).click()
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, locator.xpath_patient_Info_Link)))
                driver.find_element_by_xpath(locator.xpath_patient_Info_Link).click()
                start_time = time.perf_counter()
                sf.ajax_preloader_wait(driver)
                total_time = time.perf_counter() - start_time
                current_url = driver.current_url
                access_message = sf.CheckAccessDenied(current_url)

                if access_message == 1:
                    print("Access Denied found!")
                    # logger.critical("Access Denied found!")
                    test_case_id += 1
                    ws.append(
                        (test_case_id, patient_id, 'Navigation to Patient Info page',
                         'Failed', total_time, 'Access Denied'))

                else:
                    print("Access Check done!")
                    # logger.info("Access Check done!")
                    error_message = sf.CheckErrorMessage(driver)

                    if error_message == 1:
                        print("Error toast message is displayed")
                        # logger.critical("ERROR TOAST MESSAGE IS DISPLAYED!")
                        test_case_id += 1
                        ws.append((test_case_id,
                                   patient_id, 'Navigation to Patient Info page',
                                   'Failed', total_time, 'Error toast message is displayed'))

                    else:
                        """ **** COVERAGE **** """
                        WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located
                            ((By.XPATH, locator.xpath_patient_Info_Coverage_Link)))
                        driver.find_element_by_xpath(locator.xpath_patient_Info_Coverage_Link).click()
                        start_time = time.perf_counter()
                        sf.ajax_preloader_wait(driver)
                        total_time = time.perf_counter() - start_time
                        time.sleep(1)
                        coverage_number = len \
                            (driver.find_elements_by_xpath("//table[@id='patient_info_payment_table']"))
                        if coverage_number != 0:
                            test_case_id += 1
                            ws.append((test_case_id, patient_id, "Patient Info-->Coverage", 'Passed', total_time,
                                       'Number of Coverage card(s): ' + str(coverage_number)))
                        elif coverage_number == 0:
                            test_case_id += 1
                            ws.append((test_case_id, patient_id + ": Patient Info-->Coverage", 'Failed', total_time,
                                       'Number of Coverage card(s): ' + str(coverage_number)))

                        """ **** PATIENT INFO **** """
                        WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located(
                                (By.XPATH, locator.xpath_patient_Info_Care_Team_Link)))
                        driver.find_element_by_xpath(locator.xpath_patient_Info_Care_Team_Link).click()
                        time.sleep(1)
                        careteam_provider_number = len(driver.find_elements_by_xpath("//div[@class='mlm']/div"))
                        if careteam_provider_number != 0:
                            test_case_id += 1
                            ws.append((test_case_id, total_time, "Patient Info-->Care Team", 'Passed', 'x'
                                                                                                       'Number of Providers present in Care Team: ' + str
                                       (careteam_provider_number)))
                        elif coverage_number == 0:
                            test_case_id += 1
                            ws.append((test_case_id, patient_id + ": Patient Info-->Care Team", 'Failed',
                                       'Number of Providers present in Care Team: ' + str
                                       (careteam_provider_number)))
                        logger.info("Navigated to Patient Info.")

    except Exception as e:
        print(e)
        traceback.print_exc()
        traceback.print_exc()

        ws.append(['1', 'Unable to navigate to a patient from random metric'])
    finally:
        if window_switched == 1:
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

    driver.get(main_registry_url)
    rows = ws.max_row
    cols = ws.max_column
    for i in range(2, rows + 1):
        for j in range(3, cols + 1):
            if ws.cell(i, j).value == 'Passed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='0FC404')
            elif ws.cell(i, j).value == 'Failed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FC0E03')
            elif ws.cell(i, j).value == 'Data table is empty':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FCC0BB')


def provider_registry(driver, workbook, logger, run_from):
    workbook.create_sheet('Provider Registry')
    ws = workbook['Provider Registry']

    ws.append(['ID', 'Context', 'Scenario', 'Status', 'Time Taken', 'Comments'])
    header_font = Font(color='FFFFFF', bold=False, size=12)
    header_cell_color = PatternFill('solid', fgColor='030303')
    ws['A1'].font = header_font
    ws['A1'].fill = header_cell_color
    ws['B1'].font = header_font
    ws['B1'].fill = header_cell_color
    ws['C1'].font = header_font
    ws['C1'].fill = header_cell_color
    ws['D1'].font = header_font
    ws['D1'].fill = header_cell_color
    ws['E1'].font = header_font
    ws['E1'].fill = header_cell_color
    ws['F1'].font = header_font
    ws['F1'].fill = header_cell_color
    ws.name = "Arial"
    test_case_id = 1

    # checking for default context and then navigating to a random provider registry
    main_registry_url = driver.current_url
    if run_from == "Cozeva Support" or run_from == "Limited Cozeva Support" or run_from == "Customer Support" or run_from == "Regional Support" or run_from == "Office Admin Practice Delegate":
        try:
            print("1")
            driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()
            driver.find_element_by_id("providers-list").click()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "metric-support-prov-ls")))
            list_of_provider_elements = driver.find_element_by_id("metric-support-prov-ls").find_elements_by_tag_name(
                'tr')
            global global_search_prov
            global_search_prov = list_of_provider_elements[
                sf.RandomNumberGenerator(len(list_of_provider_elements), 1)[0]].find_elements_by_tag_name('a')[
                1].text
            list_of_provider_elements[
                sf.RandomNumberGenerator(len(list_of_provider_elements), 1)[0]].find_elements_by_tag_name('a')[
                1].click()
        except Exception as e:
            ws.append([test_case_id, "Attempting to navigate to a random provider", 'Navigation to provider context',
                       'Failed', 'x',
                       "Unable to navigate to a provider. Either the Provider list is unreachable or navigation access is denied"])
            test_case_id += 1
            driver.get(main_registry_url)
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
            print(e)
            traceback.print_exc()
            return

    # Store registry url for back navigation
    registry_url = driver.current_url

    # Navigation test 1 : Navigation to patient context through providers patients tab
    try:
        current_context = driver.find_element_by_class_name("current_context").text
        print('1.5')
        driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()
        print('2')
        driver.find_element_by_id("all_patients_tab").click()
        start_time = time.perf_counter()
        sf.ajax_preloader_wait(driver)
        total_time = time.perf_counter() - start_time
        current_url = driver.current_url
        access_message = sf.CheckAccessDenied(current_url)

        if access_message == 1:
            print("Access Denied found!")
            # logger.critical("Access Denied found!")
            test_case_id += 1
            ws.append((test_case_id, current_context, 'Navigation to all patients tab',
                       'Failed', 'x', 'Access Denied'))

        else:
            print("Access Check done!")
            # logger.info("Access Check done!")
            error_message = sf.CheckErrorMessage(driver)

            if error_message == 1:
                print("Error toast message is displayed")
                # logger.critical("ERROR TOAST MESSAGE IS DISPLAYED!")
                test_case_id += 1
                ws.append \
                    ((test_case_id, current_context, 'Navigation to all patients tab ',
                      'Failed', 'x', 'Error toast message is displayed'))

            else:
                test_case_id += 1
                ws.append((test_case_id, current_context, 'Navigation to all patients tab',
                           'Passed', total_time))
                logger.info(current_context + ": Navigated to all patients tab.")
                # Now navigating to a patient dashboard at random
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.ID, "all_patients")))
                patient_elements = driver.find_element_by_id("all_patients").find_element_by_tag_name(
                    'tbody').find_elements_by_tag_name('tr')
                patient_elements[sf.RandomNumberGenerator(len(patient_elements), 1)[0]].find_element_by_class_name(
                    'pat_name').click()
                driver.switch_to.window(driver.window_handles[1])
                start_time = time.perf_counter()
                sf.ajax_preloader_wait(driver)
                total_time = time.perf_counter() - start_time
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, locator.xpath_cozeva_Id)))
                patient_id = driver.find_element_by_xpath(locator.xpath_cozeva_Id).text
                current_url = driver.current_url
                access_message = sf.CheckAccessDenied(current_url)

                if access_message == 1:
                    print("Access Denied found!")
                    # logger.critical("Access Denied found!")
                    test_case_id += 1
                    ws.append((test_case_id, patient_id, 'Navigation to dashboard page',
                               'Failed', 'x', 'Access Denied'))


                else:
                    print("Access Check done!")
                    # logger.info("Access Check done!")
                    error_message = sf.CheckErrorMessage(driver)

                    if error_message == 1:
                        print("Error toast message is displayed")
                        # logger.critical("ERROR TOAST MESSAGE IS DISPLAYED!")
                        test_case_id += 1
                        ws.append \
                            ((test_case_id, patient_id, 'Navigation to dashboard page ',
                              'Failed', 'x', 'Error toast message is displayed'))

                    else:
                        measure_count_dashboard = len \
                            (driver.find_elements_by_xpath("//tbody[@class='measurement-body careops-new']/tr"))
                        test_case_id += 1
                        ws.append((test_case_id, patient_id, 'Navigation to dashboard page',
                                   'Passed', total_time,
                                   'Measures count in dashboard: ' + str(measure_count_dashboard)))
                        logger.info(patient_id + ": Navigated to patient dashboard.")
                        if sf.check_exists_by_class(driver, 'primary_val'):
                            test_case_id += 1
                            ws.append([test_case_id, patient_id, 'CareOps count present', 'Passed', 'x',
                                       'Count: ' + driver.find_element_by_class_name("primary_val").text])
                        else:
                            test_case_id += 1
                            ws.append([test_case_id, patient_id, 'CareOps count present', 'Failed', 'x',
                                       'Careops count not present'])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
        driver.get(registry_url)
        sf.ajax_preloader_wait(driver)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))

    except Exception as e:
        ws.append([test_case_id, "Provider registry navigation",
                   "Navigation to patient context through providers patients tab", 'Failed', 'x',
                   'Unable to navigate to patients list/Patient dashboard'])
        test_case_id += 1
        print(e)
        traceback.print_exc()
        driver.get(registry_url)
        sf.ajax_preloader_wait(driver)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))

    # Navigation test 2: Navigation to patient context through providers MSPL
    try:
        current_context = driver.find_element_by_class_name("current_context").text
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "registry_body")))
        metrics = driver.find_element_by_id("registry_body").find_elements_by_tag_name('li')
        percent = '0.00'
        loop_counter = 0
        while percent == '0.00' or percent == '0.00%':
            if loop_counter < 10:
                if len(metrics) > 1:
                    selectedMetric = metrics[sf.RandomNumberGenerator(len(metrics), 1)[0]]
                    percent = selectedMetric.find_element_by_class_name('percent').text
                elif len(metrics) == 1:
                    selectedMetric = metrics[0]
                    percent = selectedMetric.find_element_by_class_name('percent').text
                else:
                    ws.append(['No Metrics on this Provider Registry'])
                    raise Exception("No Metrics on this Provider Registry")

            else:
                ws.append(['Quit this section because there are no metrics with patients'])
                raise Exception("No Metrics with Available Patients")
        selectedMetric.click()
        start_time = time.perf_counter()
        sf.ajax_preloader_wait(driver)
        total_time = time.perf_counter() - start_time
        test_case_id += 1
        ws.append([test_case_id, current_context, "Navigation to MSPL", 'Passed', total_time])
        window_switched = 0
        try:
            patient_id = 'Couldn\'t Fetch'
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "quality_registry_list")))
            patients = driver.find_element_by_id("quality_registry_list").find_element_by_tag_name(
                'tbody').find_elements_by_tag_name('tr')
            if len(patients) > 1:
                patients[sf.RandomNumberGenerator(len(patients), 1)[0]].find_element_by_class_name('pat_name').click()
            elif len(patients) == 1:
                patients[0].find_element_by_class_name('pat_name').click()
            else:
                ws.append(['Mspl has no patients: ' + global_search_prov])
            #Add clause to check for no patients
            driver.switch_to.window(driver.window_handles[1])
            window_switched = 1
            start_time = time.perf_counter()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_cozeva_Id)))
            total_time = time.perf_counter() - start_time

            patient_id = driver.find_element_by_xpath(locator.xpath_cozeva_Id).text

            measure_count_dashboard = len \
                (driver.find_elements_by_xpath("//tbody[@class='measurement-body careops-new']/tr"))
            test_case_id += 1
            ws.append((test_case_id, patient_id, 'Navigation to dashboard page',
                       'Passed', total_time, 'Measures count in dashboard: ' + str(measure_count_dashboard)))
            logger.info(patient_id + ": Navigated to patient dashboard.")
            if sf.check_exists_by_class(driver, 'primary_val'):
                test_case_id += 1
                ws.append([test_case_id, patient_id, 'CareOps count present', 'Passed', 'x',
                           'Count: ' + driver.find_element_by_class_name("primary_val").text])
            else:
                test_case_id += 1
                ws.append(
                    [test_case_id, patient_id, 'CareOps count present', 'Failed', 'x', 'Careops count not present'])
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            driver.get(registry_url)
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))

        except Exception as e:
            test_case_id += 1
            print(e)
            traceback.print_exc()
            ws.append(
                [test_case_id, patient_id, 'clicking on random patient from patient list of Provider\'s MSPL', 'Failed',
                 '', 'Unable to click on a random patient from the MSPL'])
            if window_switched == 1:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
            driver.get(main_registry_url)
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))


    except Exception as e:
        ws.append([test_case_id, "Provider registry navigation", "Navigation to patient context through providers MSPL",
                   'Failed', 'x', 'Unable to navigate to patients list'])
        test_case_id += 1
        print(e)
        traceback.print_exc()
        driver.get(main_registry_url)
        sf.ajax_preloader_wait(driver)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))

    driver.get(main_registry_url)
    sf.ajax_preloader_wait(driver)
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
    rows = ws.max_row
    cols = ws.max_column
    for i in range(2, rows + 1):
        for j in range(3, cols + 1):
            if ws.cell(i, j).value == 'Passed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='0FC404')
            elif ws.cell(i, j).value == 'Failed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FC0E03')
            elif ws.cell(i, j).value == 'Data table is empty':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FCC0BB')


def practice_registry(driver, workbook, logger, run_from):
    workbook.create_sheet('Practice Registry')
    ws = workbook['Practice Registry']

    ws.append(['ID', 'Context', 'Scenario', 'Status', 'Time Taken', 'Comments'])
    header_font = Font(color='FFFFFF', bold=False, size=12)
    header_cell_color = PatternFill('solid', fgColor='030303')
    ws['A1'].font = header_font
    ws['A1'].fill = header_cell_color
    ws['B1'].font = header_font
    ws['B1'].fill = header_cell_color
    ws['C1'].font = header_font
    ws['C1'].fill = header_cell_color
    ws['D1'].font = header_font
    ws['D1'].fill = header_cell_color
    ws['E1'].font = header_font
    ws['E1'].fill = header_cell_color
    ws['F1'].font = header_font
    ws['F1'].fill = header_cell_color
    ws.name = "Arial"
    test_case_id = 1
    window_switched = 0
    print('Start Practice Registry Block')
    main_registry_url = driver.current_url
    # Checking run_froms and navigating to a practice registry at random
    if run_from == "Cozeva Support" or run_from == "Limited Cozeva Support" or run_from == "Customer Support" or run_from == "Regional Support":
        # Switching to random Practice name from default set context, main page
        try:
            driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()
            driver.find_element_by_id("providers-list").click()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'tabs')))
            driver.find_element_by_class_name("tabs").find_elements_by_tag_name('li')[0].click()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "metric-support-prac-ls")))
            list_of_practice_elements = driver.find_element_by_id("metric-support-prac-ls").find_elements_by_tag_name(
                'tr')
            global global_search_prac
            global_search_prac = list_of_practice_elements[
                sf.RandomNumberGenerator(len(list_of_practice_elements), 1)[0]].find_element_by_tag_name('a').text

            list_of_practice_elements[
                sf.RandomNumberGenerator(len(list_of_practice_elements), 1)[0]].find_element_by_tag_name('a').click()
            sf.ajax_preloader_wait(driver)
        except Exception as e:
            ws.append(['1', "Attempting to navigate to a random practice", 'Navigation to practice context', 'Failed',
                       "Unable to navigate to a practice. Either the Practice list is unreachable or navigation access is denied"])
            driver.get(main_registry_url)
            print(e)
            traceback.print_exc()
            sf.ajax_preloader_wait(driver)
            return
    elif run_from == "Office Admin Provider Delegate" or run_from == "Provider":
        ws.append(["1", run_from + " Role does not have access to practice Submenus"])
        return
    context_name = "Couldn't Fetch"
    registry_url = driver.current_url
    # Nav check one : Navigation to provider registry through MSPL of a practice
    try:
        # selecting a random non zero metric from the registry
        context_name = driver.find_element_by_class_name("specific_most").text
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "registry_body")))
        metrics = driver.find_element_by_id("registry_body").find_elements_by_tag_name('li')
        percent = '0.00'
        loop_count = 0
        while percent == '0.00' or percent == '0.00%':
            if loop_count < 10:
                selectedMetric = metrics[sf.RandomNumberGenerator(len(metrics), 1)[0]]
                percent = selectedMetric.find_element_by_class_name('percent').text
                loop_count += 1
            else:
                ws.append([test_case_id, context_name, 'Skipped this because control was stuck in infinite loop'])
                driver.get(main_registry_url)
                sf.ajax_preloader_wait(driver)
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
                return
        selected_metric_name = selectedMetric.find_element_by_class_name('met-name').text
        selectedMetric.click()
        sf.ajax_preloader_wait(driver)

        # clicking on a random provider name from the practice MSPL
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located(
                    (By.ID, 'metric-support-prov-ls')))
            providers = driver.find_element_by_id("metric-support-prov-ls").find_element_by_tag_name(
                'tbody').find_elements_by_tag_name('tr')
            if len(providers) != 1:
                selected_provider = \
                    providers[sf.RandomNumberGenerator(len(providers), 1)[0]].find_elements_by_tag_name('a')[2]
                selected_provider_name = selected_provider.text
                selected_provider.click()
            else:
                selected_provider_name = providers[0].find_elements_by_tag_name('a')[2].text
                providers[0].find_elements_by_tag_name('a')[2].click()
            start_time = time.perf_counter()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
            time_taken = round((time.perf_counter() - start_time), 3)
            if len(driver.find_elements_by_xpath(locator.xpath_filter_measure_list)) != 0:
                ws.append([test_case_id, selected_provider_name,
                           "Navigation to provider registry through MSPL of a practice: " + selected_metric_name,
                           'Passed', time_taken])
                test_case_id += 1
                driver.get(registry_url)

            else:
                ws.append(
                    [test_case_id, selected_provider_name,
                     "Navigation to provider registry through MSPL of a practice: " + selected_metric_name,
                     'Failed', time_taken])
                test_case_id += 1
                driver.get(registry_url)



        except Exception as e:
            print(e)
            traceback.print_exc()
            ws.append(
                [test_case_id, context_name, "Navigation to provider registry through MSPL of a practice", 'Failed', '',
                 'Couldn\'t navigate into a random provider from the MSPL: ' + selected_metric_name])
            test_case_id += 1
            driver.get(registry_url)

    except Exception as e:
        print(e)
        traceback.print_exc()
        ws.append(
            [test_case_id, context_name, "Navigation to provider registry through MSPL of a practice", 'Failed', '',
             'Couldn\'t navigate into a random metric from the provivdr registry'])
        test_case_id += 1
        print(driver.current_url)
        driver.get(registry_url)


    # Nav check two : Navigation to patient context through patient toggle of practice Metric Specific List

    try:
        sf.ajax_preloader_wait(driver)
        # selecting a random non zero metric from the registry
        print(driver.current_url)
        context_name = driver.find_element_by_class_name("specific_most").text
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "registry_body")))
        metrics = driver.find_element_by_id("registry_body").find_elements_by_tag_name('li')
        percent = '0.00'
        loop_count = 0
        while percent == '0.00' or percent == '0.00%':
            if loop_count < 10:
                selectedMetric = metrics[sf.RandomNumberGenerator(len(metrics), 1)[0]]
                percent = selectedMetric.find_element_by_class_name('percent').text
                loop_count += 1
            else:
                ws.append([test_case_id, context_name, 'Skipped this because control was stuck in infinite loop'])
                driver.get(main_registry_url)
                sf.ajax_preloader_wait(driver)
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
                return
        selected_metric_name = selectedMetric.find_element_by_class_name('met-name').text
        selectedMetric.click()
        sf.ajax_preloader_wait(driver)
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'tabs')))
            driver.find_element_by_class_name('tabs').find_elements_by_class_name('tab')[2].click()
            start_time = time.perf_counter()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'tabs')))
            time_taken = time.perf_counter() - start_time
            if driver.find_elements_by_id('performance_details') != 0:
                ws.append([test_case_id, context_name,
                           'Navigation to Performance Stats from Practice Metric : ' + selected_metric_name, 'Passed',
                           time_taken])
                test_case_id += 1
            else:
                ws.append([test_case_id, context_name,
                           'Navigation to Performance Stats from Practice Metric : ' + selected_metric_name, 'Failed'])
                test_case_id += 1

        except Exception as e:
            print(e)
            traceback.print_exc()
            ws.append([test_case_id, context_name, 'Navigation to Performance Stats from Practice MSPL', 'Failed', '',
                       'Couldnt click on the performance tab of metric :' + selected_metric_name])
            test_case_id += 1

        try:
            driver.find_element_by_class_name('tabs').find_elements_by_class_name('tab')[1].click()
            sf.ajax_preloader_wait(driver)
            print("Preloader gone")
            time.sleep(0.5)
            if len(driver.find_elements_by_class_name("ajax_preloader")) != 0:
                print("Preloader Reappeared")
                sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "metric-support-pat-ls")))
            patients = driver.find_element_by_id("metric-support-pat-ls").find_element_by_tag_name(
                'tbody').find_elements_by_tag_name('tr')
            selected_patient = patients[sf.RandomNumberGenerator(len(patients), 1)[0]].find_element_by_class_name(
                'pat_name')
            selected_patient_name = selected_patient.text
            selected_patient.click()
            driver.switch_to.window(driver.window_handles[1])
            window_switched = 1
            start_time = time.perf_counter()
            sf.ajax_preloader_wait(driver)
            print("Preloader gone")
            time.sleep(0.5)
            if len(driver.find_elements_by_class_name("ajax_preloader")) != 0:
                print("Preloader Reappeared")
                sf.ajax_preloader_wait(driver)
            time_taken = round((time.perf_counter() - start_time), 3)
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CLASS_NAME, "primary_val")))
            if len(driver.find_elements_by_class_name("primary_val")) != 0:
                ws.append([test_case_id, selected_patient_name,
                           "Navigation to patient context through patient toggle of practice Metric Specific List",
                           'Passed', time_taken])
                test_case_id += 1
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                driver.get(registry_url)
            else:
                ws.append([test_case_id, selected_patient_name,
                           "Navigation to patient context through patient toggle of practice Metric Specific List",
                           'Failed', time_taken])
                test_case_id += 1
                if window_switched == 1:
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                driver.get(registry_url)

        except Exception as e:
            print(e)
            traceback.print_exc()
            ws.append(
                [test_case_id, context_name,
                 "Navigation to patient context through patient toggle of practice Metric Specific List", 'Failed', '',
                 'Couldn\'t navigate into a random provider from the MSPL'])
            test_case_id += 1
            if window_switched == 1:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
            driver.get(registry_url)


    except Exception as e:
        print(e)
        traceback.print_exc()
        print(driver.current_url)
        ws.append(
            [test_case_id, context_name,
             "Navigation to patient context through patient toggle of practice Metric Specific List", 'Failed', '',
             'Couldn\'t navigate into a random metric from the provivdr registry'])
        test_case_id += 1

        driver.get(registry_url)

    # nav check 3 : Navigation to provider registry through providers tab in of a practice
    registry_url = driver.current_url
    try:
        driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()
        driver.find_element_by_id("providers-list").click()
        sf.ajax_preloader_wait(driver)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "metric-support-prov-ls")))
        list_of_provider_elements = driver.find_element_by_id("metric-support-prov-ls").find_elements_by_tag_name('tr')
        if len(list_of_provider_elements) > 1:
            selected_provider = list_of_provider_elements[
                sf.RandomNumberGenerator(len(list_of_provider_elements), 1)[0]].find_elements_by_tag_name('a')[1]
            selected_provider_name = selected_provider.text
            selected_provider.click()
        else:
            selected_provider = list_of_provider_elements[0].find_elements_by_tag_name('a')[1]
            selected_provider_name = selected_provider.text
            selected_provider.click()
        time_start = time.perf_counter()
        sf.ajax_preloader_wait(driver)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
        time_taken = time.perf_counter() - time_start
        if driver.find_elements_by_xpath(locator.xpath_filter_measure_list) != 0:
            ws.append([test_case_id, selected_provider_name,
                       'Navigation to provider registry through providers tab in of a practice', 'Passed',
                       round(time_taken, 3)])
            test_case_id += 1
            driver.get(registry_url)
        else:
            ws.append([test_case_id, selected_provider_name,
                       'Navigation to provider registry through providers tab in of a practice', 'Failed',
                       round(time_taken, 3)], 'Unable to locate filter element on provider\'s registry')
            test_case_id += 1
            driver.get(registry_url)

    except Exception as e:
        print(e)
        traceback.print_exc()
        ws.append([test_case_id, context_name, "Navigation to provider registry through providers tab in of a practice",
                   'Failed', "", 'Unable to click on providers\' tab and navigate to their registry'])
        test_case_id += 1
        driver.get(registry_url)

    # nav check 4 : Navigation to Performance Stats from Practice Metric specific list

    if run_from == "Cozeva Support" or run_from == "Limited Cozeva Support" or run_from == "Customer Support" or run_from == "Regional Support":
        driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()
        driver.find_element_by_id("home").click()
        sf.ajax_preloader_wait(driver)
    driver.get(main_registry_url)
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
    rows = ws.max_row
    cols = ws.max_column
    for i in range(2, rows + 1):
        for j in range(3, cols + 1):
            if ws.cell(i, j).value == 'Passed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='0FC404')
            elif ws.cell(i, j).value == 'Failed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FC0E03')
            elif ws.cell(i, j).value == 'Data table is empty':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FCC0BB')


def support_level(driver, workbook, logger, run_from):
    workbook.create_sheet('Support Level')
    ws = workbook['Support Level']

    ws.append(['ID', 'Context', 'Scenario', 'Status', 'Time Taken', 'Comments'])
    header_font = Font(color='FFFFFF', bold=False, size=12)
    header_cell_color = PatternFill('solid', fgColor='030303')
    ws['A1'].font = header_font
    ws['A1'].fill = header_cell_color
    ws['B1'].font = header_font
    ws['B1'].fill = header_cell_color
    ws['C1'].font = header_font
    ws['C1'].fill = header_cell_color
    ws['D1'].font = header_font
    ws['D1'].fill = header_cell_color
    ws['E1'].font = header_font
    ws['E1'].fill = header_cell_color
    ws['F1'].font = header_font
    ws['F1'].fill = header_cell_color
    ws.name = "Arial"
    test_case_id = 1

    registry_url = driver.current_url
    # Selecting tabs from Support MSPL
    context_name = "Couldn't Fetch"
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID, "registry_body")))
        selected_metric_name = 'Couldnt fetch Metric Name'
        context_name = driver.find_element_by_xpath(locator.xpath_context_Name).text

        metrics = driver.find_element_by_id("registry_body").find_elements_by_tag_name('li')
        percent = '0.00'
        while percent == '0.00' or percent == '0.00%':
            selectedMetric = metrics[sf.RandomNumberGenerator(len(metrics), 1)[0]]
            percent = selectedMetric.find_element_by_class_name('percent').text
        selected_metric_name = selectedMetric.find_element_by_class_name('met-name').text
        selectedMetric.click()
        sf.ajax_preloader_wait(driver)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'tab')))
        metric_url = driver.current_url
        # nav 1 : Practice Tab
        try:
            selectedPracticeName = 'Couldn\'t Fetch'
            driver.find_element_by_class_name('tabs').find_elements_by_class_name('tab')[0].click()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "metric-support-prac-ls")))
            practices = driver.find_element_by_id("metric-support-prac-ls").find_element_by_tag_name(
                'tbody').find_elements_by_tag_name('tr')
            if len(practices) > 1:
                selectedPractice = \
                    practices[sf.RandomNumberGenerator(len(practices), 1)[0]].find_elements_by_tag_name('a')[1]
                selectedPracticeName = selectedPractice.text
                global global_search_prac
                global_search_prac = selectedPracticeName
                selectedPractice.click()
            else:
                selectedPractice = practices[0].find_elements_by_tag_name('a')[1]
                selectedPracticeName = selectedPractice.text
                # global global_search_prac
                global_search_prac = selectedPracticeName
                selectedPractice.click()
            start_time = time.perf_counter()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
            time_taken = round((time.perf_counter() - start_time), 3)
            if len(driver.find_elements_by_xpath(locator.xpath_filter_measure_list)) != 0:
                ws.append([test_case_id, selectedPracticeName,
                           'Nagivation to practice Registry from the practice tab of support MSPL: ' + selected_metric_name,
                           'Passed', time_taken])
                test_case_id += 1
                driver.get(metric_url)
            else:
                ws.append([test_case_id, selectedPracticeName,
                           'Nagivation to practice Registry from the practice tab of support MSPL: ' + selected_metric_name,
                           'Failed', time_taken, 'Couldnt load into registry of a practice'])
                test_case_id += 1
                driver.get(metric_url)

        except Exception as e:
            ws.append([test_case_id, context_name,
                       'Navigation to a practice registry from the pratice tab of support MSPL :' + selected_metric_name,
                       'Failed', '',
                       'Couldnt click on practice tab or a random practice name: ' + selectedPracticeName])
            test_case_id += 1
            print(e)
            traceback.print_exc()
            driver.get(metric_url)

        # Nav to provider registry
        try:
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'tab')))
            selectedProviderName = 'Couldn\'t Fetch'
            driver.find_element_by_class_name('tabs').find_elements_by_class_name('tab')[1].click()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "metric-support-prov-ls")))
            providers = driver.find_element_by_id("metric-support-prov-ls").find_element_by_tag_name(
                'tbody').find_elements_by_tag_name('tr')
            if len(providers) > 1:
                selectedProvider = \
                    providers[sf.RandomNumberGenerator(len(providers), 1)[0]].find_elements_by_tag_name('a')[2]
                selectedProviderName = selectedProvider.text
                global global_search_prov
                global_search_prov = selectedProviderName
                selectedProvider.click()
            else:
                selectedProvider = providers[0].find_elements_by_tag_name('a')[2]
                selectedProviderNameName = selectedProvider.text
                # global global_search_prov
                global_search_prov = selectedProviderName
                selectedProvider.click()
            start_time = time.perf_counter()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
            time_taken = round((time.perf_counter() - start_time), 3)
            if len(driver.find_elements_by_xpath(locator.xpath_filter_measure_list)) != 0:
                ws.append([test_case_id, selectedProviderName,
                           'Nagivation to provider Registry from the provider tab of support MSPL: ' + selected_metric_name,
                           'Passed', time_taken])
                test_case_id += 1
                driver.get(metric_url)
            else:
                ws.append([test_case_id, selectedProviderName,
                           'Nagivation to provider Registry from the provider tab of support MSPL: ' + selected_metric_name,
                           'Failed', time_taken, 'Couldnt load into registry of a provider'])
                test_case_id += 1
                driver.get(metric_url)

        except Exception as e:
            print(e)
            traceback.print_exc()
            ws.append([test_case_id, context_name,
                       'Navigation to a provider registry from the provider tab of support MSPL :' + selected_metric_name,
                       'Failed', '',
                       'Couldnt click on provider tab or a random provider name: ' + selectedProviderName])
            test_case_id += 1
            driver.get(metric_url)

        # nav 3 : Patient context
        try:
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'tab')))
            patient_id = 'Couldn\'t Fetch'
            driver.find_element_by_class_name('tabs').find_elements_by_class_name('tab')[2].click()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "metric-support-pat-ls")))
            patients = driver.find_element_by_id("metric-support-pat-ls").find_element_by_tag_name(
                'tbody').find_elements_by_tag_name('tr')
            if len(patients) > 1:
                selectedPatient = \
                    patients[sf.RandomNumberGenerator(len(patients), 1)[0]].find_elements_by_class_name('pat_name')[0]
                selectedPatient.click()
            else:
                selectedPatient = patients[0].find_elements_by_class_name('pat_name')[0]
                selectedPatient.click()
            driver.switch_to.window(driver.window_handles[1])
            start_time = time.perf_counter()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_cozeva_Id)))
            time_taken = round((time.perf_counter() - start_time), 3)
            patient_id = driver.find_element_by_xpath(locator.xpath_cozeva_Id).text
            global global_search_pat
            global_search_pat = patient_id
            if sf.check_exists_by_class(driver, 'primary_val'):
                measure_count_dashboard = len(
                    driver.find_elements_by_xpath("//tbody[@class='measurement-body careops-new']/tr"))
                test_case_id += 1
                ws.append(
                    (test_case_id, patient_id, 'Navigation to patient context from the patients tab of support MSPL',
                     'Passed', time_taken, 'Measures count in dashboard: ' + str(measure_count_dashboard)))
                logger.info(patient_id + ": Navigated to patient dashboard.")
                test_case_id += 1
                ws.append(
                    [test_case_id, patient_id, 'Navigation to patient context from the patients tab of support MSPL',
                     'Passed', 'x', 'Count: ' + driver.find_element_by_class_name("primary_val").text])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                driver.get(metric_url)
            else:
                test_case_id += 1
                ws.append(
                    [test_case_id, patient_id, 'Navigation to patient context from the patients tab of support MSPL',
                     'Failed', 'x', 'Careops count not present'])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                driver.get(metric_url)


        except Exception as e:
            print(e)
            traceback.print_exc()
            ws.append([test_case_id, context_name,
                       'Navigation to patient context from the patients tab of support MSPL :' + selected_metric_name,
                       'Failed', '', 'Couldnt click on patient tab or a random patient : ' + patient_id])
            test_case_id += 1
            driver.get(metric_url)

        # nav 4 : Performance Statistics
        try:
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'tab')))
            driver.find_element_by_class_name('tabs').find_elements_by_class_name('tab')[3].click()
            start_time = time.perf_counter()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'tabs')))
            time_taken = time.perf_counter() - start_time
            if driver.find_elements_by_id('performance_details') != 0:
                ws.append([test_case_id, context_name,
                           'Navigation to Performance Stats from Support Metric : ' + selected_metric_name, 'Passed',
                           time_taken])
                test_case_id += 1
            else:
                ws.append([test_case_id, context_name,
                           'Navigation to Performance Stats from Support Metric : ' + selected_metric_name, 'Failed'])
                test_case_id += 1

        except Exception as e:
            print(e)
            traceback.print_exc()
            ws.append([test_case_id, context_name, 'Navigation to Performance Stats from Practice MSPL', 'Failed', '',
                       'Couldnt click on the performance tab of metric :' + selected_metric_name])
            test_case_id += 1

    except Exception as e:
        print(e)
        traceback.print_exc()
        ws.append([test_case_id, context_name, 'Navigation to Support MSPL', 'Failed', '',
                   'Unable to click on a random metric: ' + selected_metric_name])
        test_case_id += 1
        driver.get(registry_url)

    driver.get(registry_url)

    rows = ws.max_row
    cols = ws.max_column
    for i in range(2, rows + 1):
        for j in range(3, cols + 1):
            if ws.cell(i, j).value == 'Passed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='0FC404')
            elif ws.cell(i, j).value == 'Failed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FC0E03')
            elif ws.cell(i, j).value == 'Data table is empty':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FCC0BB')


def global_search(driver, workbook, logger, run_from):
    workbook.create_sheet('Global Search')
    ws = workbook['Global Search']
    ws.append(['ID', 'Context', 'Scenario', 'Status', 'Time Taken', 'Comments'])
    header_font = Font(color='FFFFFF', bold=False, size=12)
    header_cell_color = PatternFill('solid', fgColor='030303')
    ws['A1'].font = header_font
    ws['A1'].fill = header_cell_color
    ws['B1'].font = header_font
    ws['B1'].fill = header_cell_color
    ws['C1'].font = header_font
    ws['C1'].fill = header_cell_color
    ws['D1'].font = header_font
    ws['D1'].fill = header_cell_color
    ws['E1'].font = header_font
    ws['E1'].fill = header_cell_color
    ws.name = "Arial"
    test_case_id = 1
    main_registry_url = driver.current_url

    global global_search_prov, global_search_prac, global_search_pat

    def fetchPracName():
        return "test"

    def fetchProvName():
        return "test"

    def fetchPatId():
        return "test"

    def performPracSearch():
        try:
            window_switched = 0
            driver.find_element_by_id('globalsearch_input').send_keys(global_search_prac)
            start_time = time.perf_counter()
            WebDriverWait(driver, 45).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'collection-header')))
            time_taken = round(time.perf_counter() - start_time)
            driver.find_element_by_id('globalsearch_input').send_keys(Keys.RETURN)
            sf.ajax_preloader_wait(driver)
            time_taken = round(time.perf_counter() - start_time)
            # driver.find_element_by_id('globalsearch_input').send_keys(Keys.RETURN)
            driver.find_element_by_id('search_practices_link').click()
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, 'search_practices')))
            driver.find_element_by_id('search_practices').find_elements_by_tag_name('a')[0].click()
            driver.switch_to.window(driver.window_handles[1])
            window_switched = 1
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
            if len(driver.find_elements_by_xpath(locator.xpath_filter_measure_list)) != 0:
                ws.append([test_case_id, 'Practice', 'Context set to: ' + global_search_prac, 'Passed', time_taken])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                driver.get(main_registry_url)
            else:
                ws.append([test_case_id, 'Practice', 'Context set to: ' + global_search_prac, 'Failed', time_taken])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                driver.get(main_registry_url)

        except Exception as e:
            print(e)
            traceback.print_exc()
            if window_switched == 1:
                ws.append([test_case_id, 'Practice', 'Context set to: ' + global_search_prac, 'Failed', '',
                           'Unable to click on practice name from global search'])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                driver.get(main_registry_url)
            elif window_switched == 0:
                ws.append([test_case_id, 'Practice', 'Context set to: ' + global_search_prac, 'Failed', '',
                           'Unable to global search'])
                driver.get(main_registry_url)

    def performProvSearch():
        try:
            window_switched = 0
            driver.find_element_by_id('globalsearch_input').send_keys(global_search_prov)
            WebDriverWait(driver, 45).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'collection-header')))
            driver.find_element_by_id('globalsearch_input').send_keys(Keys.RETURN)
            start_time = time.perf_counter()
            sf.ajax_preloader_wait(driver)
            time_taken = round(time.perf_counter() - start_time)
            # driver.find_element_by_id('globalsearch_input').send_keys(Keys.RETURN)
            driver.find_element_by_id('search_providers_link').click()
            driver.find_element_by_id('search_providers').find_elements_by_tag_name('a')[0].click()
            driver.switch_to.window(driver.window_handles[1])
            window_switched = 1
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
            if len(driver.find_elements_by_xpath(locator.xpath_filter_measure_list)) != 0:
                ws.append([test_case_id, 'Provider', 'Context set to: ' + global_search_prov, 'Passed', time_taken])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                driver.get(main_registry_url)
            else:
                ws.append([test_case_id, 'Practice', 'Context set to: ' + global_search_prov, 'Failed', time_taken])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                driver.get(main_registry_url)

        except Exception as e:
            print(e)
            traceback.print_exc()
            if window_switched == 1:
                ws.append([test_case_id, 'Provider', 'Context set to: ' + global_search_prov, 'Failed', '',
                           'Unable to click on practice name from global search'])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                driver.get(main_registry_url)
            elif window_switched == 0:
                ws.append([test_case_id, 'Provider', 'Context set to: ' + global_search_prov, 'Failed', '',
                           'Unable to global search'])
                driver.get(main_registry_url)

    def performPatSearch():
        try:
            window_switched = 0
            driver.find_element_by_id('globalsearch_input').send_keys(global_search_pat)
            WebDriverWait(driver, 45).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'collection-header')))
            driver.find_element_by_id('globalsearch_input').send_keys(Keys.RETURN)
            start_time = time.perf_counter()
            sf.ajax_preloader_wait(driver)
            time_taken = round(time.perf_counter() - start_time)
            driver.find_element_by_id('search_patients_link').click()

            driver.find_element_by_id('search_patients').find_elements_by_tag_name('li')[
                1].find_element_by_css_selector("a[href*='patient_detail']").click()
            driver.switch_to.window(driver.window_handles[1])
            window_switched = 1
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_patient_Header_Dropdown_Arrow)))
            if len(driver.find_elements_by_xpath(locator.xpath_patient_Header_Dropdown_Arrow)) != 0:
                ws.append([test_case_id, 'Patient', 'Context set to: ' + global_search_pat, 'Passed', time_taken])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                driver.get(main_registry_url)
            else:
                ws.append([test_case_id, 'Patient', 'Context set to: ' + global_search_pat, 'Failed', time_taken])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                driver.get(main_registry_url)

        except Exception as e:
            print(e)
            traceback.print_exc()
            if window_switched == 1:
                ws.append([test_case_id, 'Patient', 'Context set to: ' + global_search_pat, 'Failed', '',
                           'Unable to click on practice name from global search'])
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                driver.get(main_registry_url)
            elif window_switched == 0:
                ws.append([test_case_id, 'Patient', 'Context set to: ' + global_search_pat, 'Failed', '',
                           'Unable to global search'])
                driver.get(main_registry_url)

    if run_from == "Cozeva Support" or run_from == "Limited Cozeva Support" or run_from == "Customer Support" or run_from == "Regional Support":
        # Perform global search for Practice, Provider and Patient
        # Fetch Practice, provider and patient ID
        if global_search_prac is None:
            global_search_prac = fetchPracName()
        if global_search_prov is None:
            global_search_prov = fetchProvName()
        if global_search_pat is None:
            global_search_pat = fetchPatId()

        performPracSearch()
        performProvSearch()
        performPatSearch()



    elif run_from == "Office Admin Practice Delegate":
        # perform global search for provider and patient
        # fetch Provider and Patient ID
        if global_search_prov is None:
            global_search_prov = fetchProvName()
        if global_search_pat is None:
            global_search_pat = fetchPatId()

        performProvSearch()
        performPatSearch()
    elif run_from == "Office Admin Provider Delegate" or run_from == "Provider":
        # perform global search for patients
        # Fetch PatientID
        if global_search_pat is None:
            global_search_pat = fetchPatId()

        performPatSearch()

    rows = ws.max_row
    cols = ws.max_column
    for i in range(2, rows + 1):
        for j in range(3, cols + 1):
            if ws.cell(i, j).value == 'Passed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='0FC404')
            elif ws.cell(i, j).value == 'Failed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FC0E03')
            elif ws.cell(i, j).value == 'Data table is empty':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FCC0BB')


def provider_mspl(driver, workbook, logger, run_from):
    workbook.create_sheet('Provider\'s MSPL')
    ws = workbook['Provider\'s MSPL']

    ws.append(['ID', 'Context', 'Scenario', 'Status', 'Time Taken', 'Comments'])
    header_font = Font(color='FFFFFF', bold=False, size=12)
    header_cell_color = PatternFill('solid', fgColor='030303')
    ws['A1'].font = header_font
    ws['A1'].fill = header_cell_color
    ws['B1'].font = header_font
    ws['B1'].fill = header_cell_color
    ws['C1'].font = header_font
    ws['C1'].fill = header_cell_color
    ws['D1'].font = header_font
    ws['D1'].fill = header_cell_color
    ws.name = "Arial"
    test_case_id = 1
    main_registry_url = driver.current_url
    selected_provider = "Couldn\'t Fetch"
    # check for default page and navigate to a Provider's Registry
    if run_from == "Cozeva Support" or run_from == "Limited Cozeva Support" or run_from == "Customer Support" or run_from == "Regional Support" or run_from == "Office Admin Practice Delegate":
        # Switching to random Practice name from default set context, main page
        try:
            driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()
            driver.find_element_by_id("providers-list").click()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.ID, "metric-support-prov-ls")))
            list_of_provider_elements = driver.find_element_by_id("metric-support-prov-ls").find_elements_by_tag_name(
                'tr')
            selected_provider = list_of_provider_elements[
                sf.RandomNumberGenerator(len(list_of_provider_elements), 1)[0]].find_elements_by_tag_name('a')[1]
            global global_search_prov
            global_search_prov = selected_provider.text
            selected_provider.click()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
        except Exception as e:
            ws.append(['1', "Attempting to navigate to a random provider", 'Navigation to provider context', 'Failed',
                       "Unable to navigate to a provider. Either the Provider list is unreachable or navigation access is denied"])
            print(e)
            traceback.print_exc()
            return
    provider_registry_url = driver.current_url
    if len(driver.find_elements_by_xpath(locator.xpath_filter_measure_list)) > 0:
        logger.info("Nagivated to Provider Registry")
        # Now, Select a random metric that is not 0/0
        try:
            metrics = driver.find_element_by_id("registry_body").find_elements_by_tag_name('li')
            percent = '0.00'
            loop_count = 0
            while percent == '0.00' or percent == '0.00%':
                if loop_count < 10:
                    if len(metrics) > 1:
                        selectedMetric = metrics[sf.RandomNumberGenerator(len(metrics), 1)[0]]
                        percent = selectedMetric.find_element_by_class_name('percent').text
                        loop_count += 1
                        print(loop_count)
                    elif len(metrics) == 1:
                        selectedMetric = metrics[0]
                        percent = selectedMetric.find_element_by_class_name('percent').text
                        loop_count += 1
                    else:
                        ws.append([
                            'No measures in selected provider\'s List: ' + global_search_prov + ', Please run Provider MSPL again'])
                        return

                else:
                    ws.append([
                        'Please Run provider MSPL again. Code was stuck in an infitite loop looking for a non 0/0 metric: ' + global_search_prov])
                    return
            selected_metric_name = selectedMetric.find_element_by_class_name('met-name').text
            selectedMetric.click()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'tabs')))
            mspl_link = driver.current_url
            # now, at the MSPL, Process starts.
            # Nav 1 : Performance Statistics
            try:
                driver.find_element_by_class_name('tabs').find_elements_by_class_name('tab')[1].click()
                start_time = time.perf_counter()
                sf.ajax_preloader_wait(driver)
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.CLASS_NAME, 'tabs')))
                time_taken = round((time.perf_counter() - start_time), 3)
                if driver.find_elements_by_id('performance_details') != 0:
                    ws.append([test_case_id, global_search_prov,
                               'Navigation to Performance Stats from MSPL : ' + selected_metric_name,
                               'Passed', time_taken])
                    test_case_id += 1
                else:
                    ws.append([test_case_id, global_search_prov,
                               'Navigation to Performance Stats from MSPL : ' + selected_metric_name,
                               'Failed', '', 'Performance Ribbon Missing'])
                    test_case_id += 1
            except Exception as e:
                print(e)
                traceback.print_exc()
                ws.append([test_case_id, global_search_prov,
                           'Navigation to Performance Stats from MSPL : ' + selected_metric_name,
                           'Failed', '', 'Unable to click on performance Statistics tab'])
                test_case_id += 1

            # nav 2: Navigation to Network Comparision
            try:
                driver.find_element_by_class_name('tabs').find_elements_by_class_name('tab')[2].click()
                start_time = time.perf_counter()
                sf.ajax_preloader_wait(driver)
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.CLASS_NAME, 'tabs')))
                time_taken = round((time.perf_counter() - start_time), 3)
                if driver.find_elements_by_id('network-table_info') != 0:
                    data_info = driver.find_element_by_id('network-table_info').text
                    ws.append([test_case_id, global_search_prov,
                               'Navigation to Network Comparision from MSPL : ' + selected_metric_name,
                               'Passed', time_taken, data_info])
                    test_case_id += 1
                else:
                    ws.append([test_case_id, global_search_prov,
                               'Navigation to Network Comparision from MSPL : ' + selected_metric_name,
                               'Failed', '', ''])
                    test_case_id += 1
            except Exception as e:
                print(e)
                traceback.print_exc()
                ws.append([test_case_id, global_search_prov,
                           'Navigation to Network comparision from MSPL : ' + selected_metric_name,
                           'Failed', '', 'Unable to click on Network Comparision tab'])
                test_case_id += 1

            # Mspl to basic careops count checking
            driver.get(mspl_link)
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'tabs')))
            window_switched = 0
            try:
                if len(driver.find_elements_by_class_name('dt_tag_value')) > 0:
                    driver.find_element_by_class_name('dt_tag_close').click()
                    sf.ajax_preloader_wait(driver)
                time.sleep(3)
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.ID, "quality_registry_list")))
                table = driver.find_element_by_id(
                    "quality_registry_list").find_element_by_tag_name(
                    "tbody").find_elements_by_tag_name('tr')
                if len(table) == 0:
                    ws.append(
                        [test_case_id, global_search_prov, 'Careops comparision between MSPL and Dashbaord', 'Failed',
                         '', 'MSPL is Empty for: ' + selected_metric_name])
                    return

                chosen_patient = randint(0, len(table) - 1)
                print(chosen_patient)
                table = driver.find_element_by_id(
                    "quality_registry_list").find_element_by_tag_name(
                    "tbody").find_elements_by_tag_name('tr')
                # Picks up the Patient's name from the MSPL
                MSPLname = table[chosen_patient].find_element_by_tag_name("a").text
                # print(MSPLname)
                # Locates the caregaps, quality gaps and hcc gaps
                MSPL_Caregap_count = table[chosen_patient].find_element_by_class_name("care_ops").text
                MSPL_caregap_hover = table[chosen_patient].find_element_by_class_name(
                    "care_ops").find_element_by_class_name(
                    "tooltipped").get_attribute(
                    "data-tooltip")
                MSPL_Quality_Gaps = table[chosen_patient].find_element_by_class_name(
                    "measure_abbr_group_pt").get_attribute(
                    "innerHTML")
                table = driver.find_element_by_id(
                    "quality_registry_list").find_element_by_tag_name(
                    "tbody").find_elements_by_tag_name('tr')
                MSPL_HCC_Gaps = table[chosen_patient].find_element_by_class_name(
                    "condition_abbr_group").get_attribute(
                    "innerHTML")
                # print("MSPL Caregaps = "+str(MSPL_Caregap_count))
                # print("MSPL Caregaps OnHover = "+MSPL_caregap_hover)
                # print("MSPL hover Quality Gaps = "+MSPL_Quality_Gaps)
                # print("MSPL hover HCC Gaps = "+MSPL_HCC_Gaps)
                MSPL_Quality_Gaps = MSPL_Quality_Gaps.split(',')
                MSPL_HCC_Gaps = MSPL_HCC_Gaps.split(',')
                MSPL_caregap_hover_count = str(len(MSPL_caregap_hover.split(',')))
                # print("MSPL Caregaps OnHover Count = " + str(MSPL_caregap_hover_count))

                # TEST CASE 1--------------------------------------------------------------------
                # print("TEST CASE 1 - MSPL CARE GAP VS # OF GAPS ON HOVER :", end=" ")
                # MSPL_Caregap_count = 21
                if MSPL_Caregap_count == MSPL_caregap_hover_count:
                    test_case_id += 1
                    ws.append((test_case_id,
                               str(MSPLname), "MSPL CARE GAP VS # OF GAPS ON HOVER",
                               'Passed', '',
                               "Caregap count on MSPL: " + MSPL_Caregap_count + " and, Number of caregaps present on hover: " + MSPL_caregap_hover_count + " . The Caregaps are: " + MSPL_caregap_hover))
                    # print("FAILED")))
                    # print("PASSED")
                    # print("TEST CASE 1 COMMENTS -", end=" ")
                    # print("Caregaps present : " + MSPL_caregap_hover)
                else:
                    test_case_id += 1
                    ws.append((test_case_id,
                               str(MSPLname), "MSPL CARE GAP VS # OF GAPS ON HOVER",
                               'Failed', '',
                               "Count mismatch between hover and caregaps, Caregap count on MSPL: " + MSPL_Caregap_count + " and, Number of caregaps present on hover: " + MSPL_caregap_hover_count + " . The Caregaps are: " + MSPL_caregap_hover))
                    # print("FAILED")
                    # print("TEST CASE 1 COMMENTS -", end=" ")
                    # print("Count mismatch between hover and caregaps")
                # TEST CASE 1--------------------------------------------------------------------

                for x in range(0, len(MSPL_Quality_Gaps)):
                    MSPL_Quality_Gaps[x] = MSPL_Quality_Gaps[x].strip()
                for x in range(0, len(MSPL_HCC_Gaps)):
                    MSPL_HCC_Gaps[x] = MSPL_HCC_Gaps[x].strip()
                MSPL_Quality_Gaps.sort()
                MSPL_HCC_Gaps.sort()
                if len(MSPL_HCC_Gaps) == 1:
                    if MSPL_HCC_Gaps[0] == "":
                        MSPL_HCC_Gaps.clear()
                if len(MSPL_Quality_Gaps) == 1:
                    if MSPL_Quality_Gaps[0] == "":
                        MSPL_Quality_Gaps.clear()

                # for x in MSPL_Quality_Gaps:
                #     print("Quality gap = "+x)
                # for x in MSPL_HCC_Gaps:
                #     print("HCC gap = "+x)

                # print("After List conversion lengths, Quality has "+str(len(MSPL_Quality_Gaps))+" Measures and hcc has
                # "+str(len(MSPL_HCC_Gaps))+" Measures") patientDashboard
                table = driver.find_element_by_id(
                    "quality_registry_list").find_element_by_tag_name(
                    "tbody").find_elements_by_tag_name('tr')
                table[chosen_patient].find_element_by_class_name("pat_name").click()

                # -------------------------Window Switch---------------------------
                # print("current window is "+driver.title)
                # parent_window = driver.current_window_handle
                window_switched = 0
                driver.switch_to.window(driver.window_handles[1])
                sf.ajax_preloader_wait(driver)
                WebDriverWait
                window_switched = 1
                # print("current window is "+driver.title)
                # -------------------------Window Switch---------------------------
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "primary_val")))
                Dashboard_caregap = driver.find_element_by_class_name("primary_val").text
                # print("Dashboard caregaps = "+Dashboard_caregap)
                driver.find_element_by_class_name("select-dropdown").click()
                dropdown_contents = driver.find_element_by_class_name("filter-panel").find_elements_by_tag_name(
                    "li")
                for dropdowntext in dropdown_contents:
                    if dropdowntext.text == "Non-Compliant":
                        dropdowntext.click()
                        break

                # TEST CASE 2--------------------------------------------------------------
                # print("TEST CASE 2 - MSPL CAREGAP COUNT VS DASHBOARD HEADER CAREOPS COUNT :", end=" ")
                if Dashboard_caregap == MSPL_caregap_hover_count:
                    test_case_id += 1
                    ws.append((test_case_id,
                               str(MSPLname), "MSPL CAREGAP COUNT VS PATIENT DASHBOARD HEADER CAREOPS COUNT",
                               'Passed', '', "Careop Count : " + Dashboard_caregap))
                    # print("PASSED")
                else:
                    test_case_id += 1
                    ws.append((test_case_id,
                               str(MSPLname), "MSPL CAREGAP COUNT VS PATIENT DASHBOARD HEADER CAREOPS COUNT",
                               'Failed', '', "Careop Count : " + Dashboard_caregap))
                    # print("FAILED")
                # print("TEST CASE 2 COMMENTS -", end=" ")
                # print("Careop Count : " + Dashboard_caregap)
                # TEST CASE 2-----------------------------------------------------------------
                Dashboard_quality_List = []
                Dashboard_HCC_List = []

                # Listize HCC Gaps------------------------------------------------------
                if sf.check_exists_by_id(driver, "table_4"):
                    hcctable = driver.find_element_by_id("table_4").find_elements_by_class_name(
                        "compliant_true")
                    # print("Dashboard Hcc Count = "+str(len(hcctable)))
                    for hcc_measures in hcctable:
                        hcc_abr = ""
                        for hcc_abr_text in hcc_measures.find_element_by_class_name(
                                "hcc_details").find_element_by_tag_name(
                            "span").text:
                            if hcc_abr_text != 'H' and hcc_abr_text != 'C' and hcc_abr_text != 'C':
                                hcc_abr = hcc_abr + hcc_abr_text
                        Dashboard_HCC_List.append("#" + hcc_abr.strip())
                    Dashboard_HCC_List.sort()
                    # for x in Dashboard_HCC_List:
                    #     print("Dashboard HCC = "+x)

                # Listize Quality Gaps----------------------------------------------------
                if sf.check_exists_by_id(driver, "table_1"):
                    qualitytable = driver.find_element_by_id("table_1").find_elements_by_class_name(
                        "compliant_true")
                    # print("Dashboard Quality Gap count = "+str(len(qualitytable)))
                    for quality_measures in qualitytable:
                        quality_abr = ""
                        for Quality_abr_text in quality_measures.find_element_by_class_name("tiny-text").text:
                            if Quality_abr_text == '·':
                                break
                            else:
                                quality_abr = quality_abr + Quality_abr_text
                        Dashboard_quality_List.append(quality_abr.strip())
                    Dashboard_quality_List.sort()
                    # for x in Dashboard_quality_List:
                    #     print("Dashboard Quality = "+x)

                # TEST CASE 3--------------------------------------------------------------
                # print("TEST CASE 3 - NUMBER OF MEASURES ON DASHBOARD VS MSPL CARE GAP COUNT :", end=" ")
                if MSPL_Caregap_count == str(len(Dashboard_quality_List) + len(Dashboard_HCC_List)):
                    test_case_id += 1
                    ws.append((test_case_id,
                               str(MSPLname), "NUMBER OF MEASURES ON PATIENT DASHBOARD VS MSPL CARE GAP COUNT",
                               'Passed', '', "Measures on dashboard : " + str(
                        len(Dashboard_quality_List) + len(Dashboard_HCC_List))))
                    # print("PASSED")
                    # print("TEST CASE 3 COMMENTS -", end=" ")
                    # print("Measures on dashboard : " + str(len(Dashboard_quality_List) + len(Dashboard_HCC_List)))
                else:
                    test_case_id += 1
                    ws.append((test_case_id,
                               str(MSPLname), "NUMBER OF MEASURES ON PATIENT DASHBOARD VS MSPL CARE GAP COUNT",
                               'Failed', '',
                               "Measures on dashboard : " + str(
                                   len(Dashboard_quality_List) + len(Dashboard_HCC_List)) +
                               ". MSPL caregap count :" + MSPL_Caregap_count + ". For a difference of " +
                               str(abs(int(MSPL_Caregap_count) - (
                                       len(Dashboard_quality_List) + len(Dashboard_HCC_List))))))
                    # print("FAILED")
                    # print("TEST CASE 3 COMMENTS -", end=" ")
                    # print("Measures on dashboard : " + str(
                    #     len(Dashboard_quality_List) + len(Dashboard_HCC_List)) + ". MSPL caregap count :" + str(
                    #     MSPL_Caregap_count) + ". For a difference of " + str(
                    #     abs(MSPL_Caregap_count - (len(Dashboard_quality_List) + len(Dashboard_HCC_List)))))
                # TEST CASE 3------------------------------------------------------------------
                # TEST CASE 4------------------------------------------------------------------
                # print("TEST CASE 4 - COMPARISION BETWEEN QUALITY GAPS ON MSPL HOVER AND DASHBOARD :", end=" ")
                if len(MSPL_Quality_Gaps) == len(Dashboard_quality_List):
                    flag = 0
                    for x in range(0, len(MSPL_Quality_Gaps)):
                        if MSPL_Quality_Gaps[x] != Dashboard_quality_List[x]:
                            Different_Measure = list(set(MSPL_Quality_Gaps) ^ set(Dashboard_quality_List))
                            flag = 1
                            break
                    if flag == 1:
                        test_case_id += 1
                        ws.append((test_case_id,
                                   str(MSPLname),
                                   "COMPARISION BETWEEN QUALITY GAPS ON MSPL HOVER AND PATIENT DASHBOARD",
                                   'Failed', '',
                                   str(len(MSPL_Quality_Gaps)) + " of " + str(len(Dashboard_quality_List)) +
                                   " Measures. Different Measures are " + ', '.join(
                                       map(str, Different_Measure))))
                        # print("FAILED")
                        # print("TEST CASE 4 COMMENTS -", end=" ")
                        # print(str(len(MSPL_Quality_Gaps)) + " of " + str(
                        #     len(Dashboard_quality_List)) + " Measures. Different Measures are " + ', '.join(
                        #     map(str, Different_Measure)))
                    elif flag == 0:
                        test_case_id += 1
                        ws.append((test_case_id,
                                   str(MSPLname),
                                   "COMPARISION BETWEEN QUALITY GAPS ON MSPL HOVER AND PATIENT DASHBOARD",
                                   'Passed', '',
                                   str(len(MSPL_Quality_Gaps)) + " of " + str(len(Dashboard_quality_List)) +
                                   " Measures which are " + ', '.join(map(str, Dashboard_quality_List))))
                        # print("PASSED")
                        Different_Measure = list(set(MSPL_Quality_Gaps) ^ set(Dashboard_quality_List))
                        # print("TEST CASE 4 COMMENTS -", end=" ")
                        # print(str(len(MSPL_Quality_Gaps)) + " of " + str(
                        #     len(Dashboard_quality_List)) + " Measures which are " + ', '.join(
                        #     map(str, Dashboard_quality_List)))
                else:
                    Different_Measure = list(set(MSPL_Quality_Gaps) ^ set(Dashboard_quality_List))
                    test_case_id += 1
                    ws.append((test_case_id,
                               str(MSPLname), "COMPARISION BETWEEN QUALITY GAPS ON MSPL HOVER AND PATIENT DASHBOARD",
                               'Failed', '',
                               str(len(MSPL_Quality_Gaps)) + " of " + str(len(Dashboard_quality_List)) +
                               " Measures. Different Measures are " + ', '.join(map(str, Different_Measure))))
                    # print("FAILED")
                    # print("TEST CASE 4 COMMENTS -", end=" ")
                    # print(str(len(MSPL_Quality_Gaps)) + " of " + str(len(Dashboard_quality_List))
                    #       + " Measures. Different Measures are " + ', '.join(map(str, Different_Measure)))
                # TEST CASE 4------------------------------------------------------------------
                # Dashboard_HCC_List.append("#45")
                # TEST CASE 5------------------------------------------------------------------
                # print("TEST CASE 5 - COMPARISION BETWEEN HCC GAPS ON MSPL HOVER AND DASHBOARD :", end=" ")
                if len(MSPL_HCC_Gaps) == len(Dashboard_HCC_List):
                    flag = 0
                    for x in range(0, len(MSPL_HCC_Gaps)):
                        if MSPL_HCC_Gaps[x] != Dashboard_HCC_List[x]:
                            Different_Measure = list(set(MSPL_HCC_Gaps) ^ set(Dashboard_HCC_List))
                            flag = 1
                            break
                    if flag == 1:
                        test_case_id += 1
                        ws.append((test_case_id,
                                   str(MSPLname), "COMPARISION BETWEEN HCC GAPS ON MSPL HOVER AND PATIENT DASHBOARD",
                                   'Failed', '', str(len(MSPL_HCC_Gaps)) + " of " + str(len(Dashboard_HCC_List)) +
                                   " Measures. Different Measures are " + ', '.join(
                            map(str, Different_Measure))))
                        # print("FAILED")
                        # print("TEST CASE 5 COMMENTS -", end=" ")
                        # print(str(len(MSPL_HCC_Gaps)) + " of " + str(
                        #     len(Dashboard_HCC_List)) + " Measures. Different Measures are " + ', '.join(
                        #     map(str, Different_Measure)))
                    elif flag == 0 & len(MSPL_HCC_Gaps) == 0:
                        test_case_id += 1
                        ws.append((test_case_id,
                                   str(MSPLname), "COMPARISION BETWEEN HCC GAPS ON MSPL HOVER AND PATIENT DASHBOARD",
                                   'Passed', '', str(len(MSPL_HCC_Gaps)) + " of " +
                                   str(len(Dashboard_HCC_List)) + " Measures"))
                        # print("PASSED")
                        # print("TEST CASE 5 COMMENTS -", end=" ")
                        # print(str(len(MSPL_HCC_Gaps)) + " of " + str(
                        #     len(Dashboard_HCC_List)) + " Measures")
                    elif flag == 0:
                        test_case_id += 1
                        ws.append((test_case_id,
                                   str(MSPLname), "COMPARISION BETWEEN HCC GAPS ON MSPL HOVER AND PATIENT DASHBOARD",
                                   'Passed', '', str(len(MSPL_HCC_Gaps)) + " of " + str(len(Dashboard_HCC_List)) +
                                   " Measures which are " + ', '.join(map(str, Dashboard_HCC_List))))
                        # print("PASSED")
                        # print("TEST CASE 5 COMMENTS -", end=" ")
                        # print(str(len(MSPL_HCC_Gaps)) + " of " + str(
                        #     len(Dashboard_HCC_List)) + " Measures which are " + ', '.join(
                        #     map(str, Dashboard_HCC_List)))
                else:
                    Different_Measure = list(set(MSPL_HCC_Gaps) ^ set(Dashboard_HCC_List))
                    test_case_id += 1
                    ws.append((test_case_id,
                               str(MSPLname), "COMPARISION BETWEEN HCC GAPS ON MSPL HOVER AND PATIENT DASHBOARD",
                               'Failed', '', str(len(MSPL_HCC_Gaps)) + " of " + str(len(Dashboard_HCC_List)) +
                               " Measures. Different Measures are " + ', '.join(map(str, Different_Measure))))
                    # print("FAILED")
                    Different_Measure = list(set(MSPL_HCC_Gaps) ^ set(Dashboard_HCC_List))
                    # print("TEST CASE 5 COMMENTS -", end=" ")
                    # print(str(len(MSPL_HCC_Gaps)) + " of " + str(
                    #     len(Dashboard_HCC_List)) + " Measures. Different Measures are " + ', '.join(
                    #     map(str, Different_Measure)))
                # TEST CASE 5------------------------------------------------------------------

                # -------------------------Window Switch---------------------------
                # print("current window is "+driver.title)
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                window_switched=0
                sf.ajax_preloader_wait(driver)
                # print("current window is "+driver.title)
                # -------------------------Window Switch---------------------------
                # ASPYEDIT END-----------------------------------------------------------------------------------------------



            except Exception as e:
                print(e)
                traceback.print_exc()
                ws.append([test_case_id, global_search_prov, 'Careops comparision from mspl: ' + selected_metric_name,
                           'Failed', '', ])
                if window_switched == 1:
                    driver.switch_to.window(driver.window_handles[0])
                    sf.ajax_preloader_wait(driver)
                    window_switched = 0

        except Exception as e:
            print(e)
            traceback.print_exc()
            logger.critical("Unable to click on a metric from the provider dashboard")
            ws.append(
                [test_case_id, selected_provider, "Attempting to click on metric in dashboard: " + selected_metric_name,
                 'Failed', '', 'Unable to click on metric'])
            test_case_id += 1
            if window_switched == 1:
                driver.switch_to.window(driver.window_handles[0])
                sf.ajax_preloader_wait(driver)
                window_switched = 0

    else:
        logger.info("Provider Registry Navigation Failed")

    driver.get(main_registry_url)
    sf.ajax_preloader_wait(driver)
    WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))


    rows = ws.max_row
    cols = ws.max_column
    for i in range(2, rows + 1):
        for j in range(3, cols + 1):
            if ws.cell(i, j).value == 'Passed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='0FC404')
            elif ws.cell(i, j).value == 'Failed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FC0E03')
            elif ws.cell(i, j).value == 'Data table is empty':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FCC0BB')


def time_capsule(driver, workbook, logger, run_from):
    workbook.create_sheet('Time Capsule')
    ws = workbook['Time Capsule']

    ws.append(['ID', 'Context', 'Scenario', 'Status', 'Comments'])
    header_font = Font(color='FFFFFF', bold=False, size=12)
    header_cell_color = PatternFill('solid', fgColor='030303')
    ws['A1'].font = header_font
    ws['A1'].fill = header_cell_color
    ws['B1'].font = header_font
    ws['B1'].fill = header_cell_color
    ws['C1'].font = header_font
    ws['C1'].fill = header_cell_color
    ws['D1'].font = header_font
    ws['D1'].fill = header_cell_color
    ws['E1'].font = header_font
    ws['E1'].fill = header_cell_color

    ws.name = "Arial"
    test_case_id = 1
    try:
        last_url = driver.current_url
        window_switched = 0
        driver.find_element_by_xpath(locator.xpath_app_Tray_Link).click()
        driver.find_element_by_xpath(locator.xpath_app_Time_Capsule).click()
        driver.switch_to.window(driver.window_handles[1])
        window_switched = 1
        sf.ajax_preloader_wait(driver)
        try:
            sf.ajax_preloader_wait(driver)
            current_url = driver.current_url
            access_message = sf.CheckAccessDenied(current_url)

            if access_message == 1:
                print("Access Denied found!")
                # logger.critical("Access Denied found!")
                test_case_id+= 1
                ws.append(
                    (test_case_id, 'Time Capsule', 'Access check for Time Capsule', 'Failed', 'Access Denied'))

            else:
                print("Access Check done!")
                # logger.info("Access Check done!")
                error_message = sf.CheckErrorMessage(driver)

                if error_message == 1:
                    print("Error toast message is displayed")
                    # logger.critical("ERROR TOAST MESSAGE IS DISPLAYED!")
                    test_case_id += 1
                    ws.append((test_case_id, 'Time Capsule', 'Navigation to Time Capsule without error message',
                                'Failed', 'Error toast message is displayed'))

                else:
                    if len(driver.find_elements_by_xpath(locator.xpath_latest_Card_Title)) != 0:
                        latest_computation_dete = driver.find_element_by_xpath(
                            locator.xpath_latest_Card_Title).text
                        test_case_id += 1
                        ws.append((test_case_id, 'Time Capsule', 'Navigation to Time Capsule',
                                    'Passed', 'Latest Computation date: ' + latest_computation_dete))
                    else:
                        test_case_id += 1
                        ws.append((test_case_id, 'Time Capsule', 'Navigation to Time Capsule', 'Failed',
                                    'Computation card details is not found!'))

        except Exception as e:
            print(e)
            test_case_id += 1
            ws.append(
                (test_case_id, 'Time Capsule', 'Navigation to Time Capsule', 'Failed', 'Exception occurred!'))
        finally:
            driver.close()
            time.sleep(1)
            if window_switched == 1:
                driver.switch_to.window(driver.window_handles[0])

    except Exception as e:
        print(e)
        test_case_id += 1
        ws.append(
            (test_case_id, 'Time Capsule', 'Navigation to Time Capsule', 'Failed', 'Exception occurred!'))
        driver.get(last_url)
        sf.ajax_preloader_wait(driver)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, locator.xpath_app_Tray_Link)))

    rows = ws.max_row
    cols = ws.max_column
    for i in range(2, rows + 1):
        for j in range(3, cols + 1):
            if ws.cell(i, j).value == 'Passed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='0FC404')
            elif ws.cell(i, j).value == 'Failed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FC0E03')
            elif ws.cell(i, j).value == 'Data table is empty':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FCC0BB')


def secure_messaging(driver, workbook, logger, run_from):
    workbook.create_sheet('Secure Messaging')
    ws = workbook['Secure Messaging']

    ws.append(['ID', 'Context', 'Scenario', 'Status', 'Comments'])
    header_font = Font(color='FFFFFF', bold=False, size=12)
    header_cell_color = PatternFill('solid', fgColor='030303')
    ws['A1'].font = header_font
    ws['A1'].fill = header_cell_color
    ws['B1'].font = header_font
    ws['B1'].fill = header_cell_color
    ws['C1'].font = header_font
    ws['C1'].fill = header_cell_color
    ws['D1'].font = header_font
    ws['D1'].fill = header_cell_color
    ws['E1'].font = header_font
    ws['E1'].fill = header_cell_color
    test_case_ID = 0

    ws.name = "Arial"

    window_switched = 0
    try:
        last_url = driver.current_url
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, locator.xpath_app_Tray_Link)))
        driver.find_element_by_xpath(locator.xpath_app_Tray_Link).click()
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, locator.xpath_app_Secure_Messaging)))
        driver.find_element_by_xpath(locator.xpath_app_Secure_Messaging).click()
        driver.switch_to.window(driver.window_handles[1])
        window_switched = 1
        try:
            sf.ajax_preloader_wait(driver)
            current_url = driver.current_url
            access_message = sf.CheckAccessDenied(current_url)

            if access_message == 1:
                print("Access Denied found!")
                # logger.critical("Access Denied found!")
                test_case_ID += 1
                ws.append(
                    (test_case_ID, 'Secure Messaging', 'Access check for Secure Messaging', 'Failed',
                     'Access Denied'))

            else:
                print("Access Check done!")
                # logger.info("Access Check done!")
                error_message = sf.CheckErrorMessage(driver)

                if error_message == 1:
                    print("Error toast message is displayed")
                    # logger.critical("ERROR TOAST MESSAGE IS DISPLAYED!")
                    test_case_ID += 1
                    ws.append(
                        (test_case_ID, 'Secure Messaging',
                         'Navigation to Secure Messaging without error message',
                         'Failed', 'Error toast message is displayed'))

                else:
                    total_inbox_messages = len(driver.find_elements_by_xpath(locator.xpath_inbox_Message))
                    test_case_ID += 1
                    ws.append((test_case_ID, 'Secure Messaging', 'Navigation to Secure Messaging', 'Passed',
                                '[Inbox]Number of messages in the first page: ' + str(total_inbox_messages)))
        except Exception as e:
            print(e)
            test_case_ID += 1
            ws.append((test_case_ID, 'Secure Messaging', 'Navigation to Secure Messaging', 'Failed',
                        'Exception occurred!'))
        finally:
            driver.close()
            time.sleep(1)
            driver.switch_to.window(driver.window_handles[0])

    except Exception as e:
        print(e)
        test_case_ID += 1
        ws.append((test_case_ID, 'Secure Messaging', 'Navigation to Secure Messaging', 'Failed',
                    'Exception occurred!'))
        driver.get(last_url)
        sf.ajax_preloader_wait(driver)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, locator.xpath_app_Tray_Link)))


    rows = ws.max_row
    cols = ws.max_column
    for i in range(2, rows + 1):
        for j in range(3, cols + 1):
            if ws.cell(i, j).value == 'Passed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='0FC404')
            elif ws.cell(i, j).value == 'Failed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FC0E03')
            elif ws.cell(i, j).value == 'Data table is empty':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FCC0BB')


def analytics(driver, workbook, logger, run_from):
    workbook.create_sheet('Analytics')
    ws = workbook['Analytics']

    ws.append(['ID', 'Context', 'Scenario', 'Status', 'Time Taken', 'Comments'])
    header_font = Font(color='FFFFFF', bold=False, size=12)
    header_cell_color = PatternFill('solid', fgColor='030303')
    ws['A1'].font = header_font
    ws['A1'].fill = header_cell_color
    ws['B1'].font = header_font
    ws['B1'].fill = header_cell_color
    ws['C1'].font = header_font
    ws['C1'].fill = header_cell_color
    ws['D1'].font = header_font
    ws['D1'].fill = header_cell_color
    ws['E1'].font = header_font
    ws['E1'].fill = header_cell_color
    ws['F1'].font = header_font
    ws['F1'].fill = header_cell_color
    ws.name = "Arial"
    test_case_id = 1
    last_url = driver.current_url

    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, locator.xpath_app_Tray_Link)))
        driver.find_element_by_xpath(locator.xpath_app_Tray_Link).click()
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, locator.xpath_app_Analytics)))
        driver.find_element_by_xpath(locator.xpath_app_Analytics).click()
        driver.switch_to.window(driver.window_handles[1])
        try:
            sf.ajax_preloader_wait(driver)
            current_url = driver.current_url
            access_message = sf.CheckAccessDenied(current_url)

            if access_message == 1:
                print("Access Denied found!")
                # logger.critical("Access Denied found!")
                test_case_id += 1
                ws.append(
                    (test_case_id, 'Analytics', 'Access check for Analytics', 'Failed', 'x', 'Access Denied'))

            else:
                print("Access Check done!")
                # logger.info("Access Check done!")
                error_message = sf.CheckErrorMessage(driver)

                if error_message == 1:
                    print("Error toast message is displayed")
                    # logger.critical("ERROR TOAST MESSAGE IS DISPLAYED!")
                    test_case_id += 1
                    ws.append((test_case_id, 'Analytics', 'Navigation to Analytics without error message',
                               'Failed', 'x', 'Error toast message is displayed'))

                else:

                    total_workbooks = len(driver.find_elements_by_xpath(locator.xpath_total_Workbooks))
                    all_workbooks = driver.find_elements_by_xpath(locator.xpath_total_Workbooks)
                    test_case_id += 1
                    ws.append((test_case_id, 'Analytics', 'Navigation to Analytics', 'Passed', 'x',
                               'Number of Workbook links: ' + str(total_workbooks)))
                    workbook_link = 0
                    while workbook_link < len(all_workbooks):
                        WebDriverWait(driver, 60).until(
                            EC.presence_of_element_located((By.XPATH, "//tr[@worksheet_title='Quality Overview']")))
                        workbook_name = (all_workbooks[workbook_link]).text
                        all_workbooks[workbook_link].click()
                        start_time = time.perf_counter()
                        print(workbook_name)
                        try:
                            WebDriverWait(driver, 100).until(EC.invisibility_of_element_located(
                                (By.XPATH, "// div[@class ='sm_download_cssload_loader']")))
                            WebDriverWait(driver, 30).until(
                                EC.presence_of_element_located((By.XPATH, "//a[@id='sm_back']")))
                            time_taken = time.perf_counter() - start_time
                            if len(driver.find_elements_by_xpath("//div[@class='nodata']")) == 0:
                                print(workbook_name + "Passed")
                                test_case_id += 1
                                ws.append((test_case_id, 'Analytics Workbook', workbook_name, 'Passed', time_taken, ''))
                            elif len(driver.find_elements_by_xpath("//div[@class='nodata']")) != 0:
                                test_case_id += 1
                                ws.append((test_case_id, 'Analytics Workbook', workbook_name, 'Failed', time_taken,
                                           'No data for the selected filters'))
                            # ASPY EDIT -------------------------------------------------------------------------------------------------------------------------
                            # '''
                            loader_element = 'sm_download_cssload_loader_wrap'
                            loader_element2 = 'toast sm_small_toast_message'
                            WebDriverWait(driver, 100).until(
                                EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
                            WebDriverWait(driver, 100).until(
                                EC.invisibility_of_element_located((By.CLASS_NAME, loader_element2)))
                            # WebDriverWait(driver, 30).until(
                            # EC.element_to_be_clickable((By.XPATH, "// *[ @ id = 'sm_select_all'] / i")))
                            time.sleep(0.5)
                            select_all_present = True
                            try:
                                driver.find_element_by_xpath("// *[ @ id = 'sm_select_all'] / i").click()
                            except Exception as e:
                                print("No select all checkbox")
                                select_all_present = False
                            if select_all_present:
                                Drilldown_links = driver.find_element_by_class_name("breadcrumb_dropdown"). \
                                    find_elements_by_tag_name("a")
                            print("hello???")
                            # Next_drilldown_present = True
                            Drilldown_links.pop(0)
                            for link in Drilldown_links:
                                if select_all_present:
                                    link.click()
                                    start_time = time.perf_counter()
                                    WebDriverWait(driver, 500).until(
                                        EC.invisibility_of_element_located((By.CLASS_NAME, loader_element)))
                                    time_taken = time.perf_counter() - start_time
                                    WebDriverWait(driver, 100).until(
                                        EC.invisibility_of_element_located((By.CLASS_NAME, loader_element2)))
                                    Worksheet_name = link.text
                                    if len(driver.find_elements_by_xpath("//div[@class='nodata']")) == 0:
                                        print(Worksheet_name + "Passed")
                                        test_case_id += 1
                                        ws.append(
                                            (test_case_id, 'Analytics Worksheet', Worksheet_name, 'Passed',
                                             (str)(round(time_taken, 3))))
                                    elif len(driver.find_elements_by_xpath("//div[@class='nodata']")) != 0:
                                        test_case_id += 1
                                        ws.append((test_case_id, 'Analytics Worksheet',
                                                   workbook_name + "-" + Worksheet_name, 'Failed',
                                                   (str)(round(time_taken, 3)),
                                                   'No data for the selected filters'))
                                    try:
                                        driver.find_element_by_xpath("// *[ @ id = 'sm_select_all'] / i").click()
                                        print("Found Select all")
                                    except Exception as e:
                                        print("No select all checkbox")
                                        print(e)
                                        select_all_present = False
                                        break
                                        # '''
                            # Aspyedit ends here ---------------------------------------------------------------------------------------------------------------------------------------
                            driver.find_element_by_xpath("//a[@id='sm_back']").click()

                        except Exception as e:
                            print(e)
                            traceback.print_exc()
                            print(workbook_name + "Failed!Exception occurred!")
                            test_case_id += 1
                            ws.append((test_case_id, 'Analytics Workbook', workbook_name, 'Failed', '', ''))
                            driver.get(current_url)

                        finally:
                            workbook_link += 1
                            all_workbooks = driver.find_elements_by_xpath(locator.xpath_total_Workbooks)

        except Exception as e:
            print(e)
            traceback.print_exc()
            test_case_id += 1
            ws.append((test_case_id, 'Analytics', 'Navigation to Analytics', 'Failed', '', 'Exception occurred!'))
        finally:
            driver.close()
            time.sleep(1)
            driver.switch_to.window(driver.window_handles[0])


    except Exception as e:
        print(e)
        traceback.print_exc()
        test_case_id += 1
        ws.append((test_case_id, 'Analytics', 'Navigation to Analytics', 'Failed', 'Exception occurred!'))
        driver.get(last_url)
        sf.ajax_preloader_wait(driver)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, locator.xpath_app_Tray_Link)))

    rows = ws.max_row
    cols = ws.max_column
    for i in range(2, rows + 1):
        for j in range(3, cols + 1):
            if ws.cell(i, j).value == 'Passed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='0FC404')
            elif ws.cell(i, j).value == 'Failed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FC0E03')
            elif ws.cell(i, j).value == 'Data table is empty':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FCC0BB')


def click_on_each_metric(customer, driver, workbook, path):
    ws = workbook.create_sheet(customer)
    ws = workbook[customer]
    ws.append(['ID', 'List', 'Context', 'Time-Taken'])
    workbook.save(path + "\\Report.xlsx")

    metrics = driver.find_element_by_id("registry_body").find_elements_by_tag_name('li')
    tracker = 1
    i = 0
    while i < len(metrics):
        print(metrics[i].text)
        try:
            driver.execute_script("arguments[0].scrollIntoView();", metrics[i])
            metrics[i].click()
            start = time.perf_counter()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 100).until(
                EC.presence_of_element_located((By.CLASS_NAME, "tabs")))
            total_time = time.perf_counter() - start
            # print('Getting Here')
            context = driver.find_element_by_class_name("metric_specific_patient_list_title").text
            print("getting here")
            ws.append([tracker, 'Provider\'s Tab', context, total_time])
            clicky = driver.find_element_by_class_name('tabs').find_elements_by_tag_name('li')
            print(*clicky)
            clicky[0].click()
            start = time.perf_counter()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 100).until(
                EC.presence_of_element_located((By.CLASS_NAME, "tabs")))
            total_time = time.perf_counter() - start
            ws.append([tracker, 'Practice Tab', context, total_time])
            print("Practice done")
            clicky = driver.find_element_by_class_name('tabs').find_elements_by_tag_name('li')
            print("found tabs from practice tab")
            clicky[2].click()
            print("Patients clicked")
            start = time.perf_counter()
            sf.ajax_preloader_wait(driver)
            WebDriverWait(driver, 100).until(
                EC.presence_of_element_located((By.CLASS_NAME, "tabs")))
            total_time = time.perf_counter() - start
            ws.append([tracker, 'Patients Tab', context, total_time])
            workbook.save(path + "\\Report.xlsx")
            driver.find_element_by_class_name("breadcrumb").click()
            print('Back clicked')
            WebDriverWait(driver, 100).until(
                EC.presence_of_element_located((By.XPATH, locator.side_nav_SlideOut)))
            print('Sidenav found')
            sf.ajax_preloader_wait(driver)
            print("Back to registries")
            i = i + 1
            metrics = driver.find_element_by_id("registry_body").find_elements_by_tag_name('li')


        except Exception as e:
            ws.append([tracker, 'Failed', 'Failed'])
            print(e)
            traceback.print_exc()
        finally:
            tracker = tracker + 1

    driver.find_element_by_xpath(locator.side_nav_SlideOut).click()
    driver.find_element_by_id("providers-list").click()
    start = time.perf_counter()
    sf.ajax_preloader_wait(driver)
    WebDriverWait(driver, 100).until(
        EC.presence_of_element_located((By.ID, "qt-mt-support-ls")))
    total_time = time.perf_counter() - start
    ws.append([tracker, 'Provider\'s List', total_time])
    tracker = tracker + 1
    driver.find_elements_by_class_name("handler")[0].click()
    start = time.perf_counter()
    sf.ajax_preloader_wait(driver)
    WebDriverWait(driver, 100).until(
        EC.presence_of_element_located((By.ID, "qt-mt-support-ls")))
    total_time = time.perf_counter() - start
    ws.append([tracker, 'Practice List', '', total_time])
    workbook.save(path + "\\Report.xlsx")


'''
def SupportpageAccordionValidationx(driver, workbook, logger, run_from):
    try:
        loader = WebDriverWait(driver, 300)
        loader.until(EC.invisibility_of_element_located((By.XPATH, "//div/div[contains(@class,'ajax_preloader')]")))
        LOBdropdownelement = driver.find_element_by_xpath("//*[@id='qt-filter-label']")
        LOBdropdownelement.click()
        default_quarter = driver.find_element_by_xpath(
            "//*[@id='filter-quarter']//following-sibling::li[@class='highlight']/span/a")
        print(default_quarter.text)
        logger.info("Default Quarter selected as  : " + default_quarter.text)
        logger.info("--------------------------------------------------------")
        Flag = config['SelectMeasurementYear']['Flag_Support']
        # Navigation with customize Measurement Year
        # LOBquarterlist = []
        print(Flag)
        LOBquarter = LOBdropdownelement.find_elements_by_xpath("//*[@id='filter-quarter']/li")
        if (Flag == "True"):
            for i in range(0, len(LOBquarter)):
                # LOBquarterlist.append(LOBquarter[i].text)
                if (LOBquarter[i].text == config['SelectMeasurementYear']['MeasurementYear_Support'] or LOBquarter[
                    i].text == config['SelectMeasurementYear']['MeasurementYearQuarter_Support']):
                    logger.info("Current Quarter selected as  : " + LOBquarter[i].text)
                    logger.info("--------------------------------------------------------")
                    LOBquarter[i].click()
                    break
                LOBquarter = LOBdropdownelement.find_elements_by_xpath("//*[@id='filter-quarter']/li")

        time.sleep(1)

        LOBname = LOBdropdownelement.find_element_by_xpath("//*[@id='filter-lob']")
        LOBnamelist = LOBname.find_elements_by_tag_name("li")
        Payername = LOBdropdownelement.find_elements_by_xpath("//*[@id='filter-payer']")
        for j in range(0, len(LOBnamelist)):

            print(LOBnamelist[j].text)
            print("--------------------------------")
            LOBnamelist[j].click()
            currentLOBName = LOBnamelist[j].text
            logger.info("LOB Selected  :: " + currentLOBName)
            logger.info("---------------------------------------")
            driver.find_element_by_xpath("//*[@id='reg-filter-apply']").click()
            time.sleep(2)
            loader = WebDriverWait(driver, 300)
            loader.until(
                EC.invisibility_of_element_located((By.XPATH, "//div/div[contains(@class,'ajax_preloader')]")))
            logger.captureScreenshot(driver, currentLOBName, screenshot_path)
            # Checking Patient count in Lob. Raise error if it is 0
            Lob_type = ["ALL", "Medicare", "Medicare ACO"]
            try:
                if currentLOBName in Lob_type:
                    lobpatientcount = driver.find_element_by_xpath(
                        "//*[@id='quality_registry']/div/div[1]/div[4]/div[2]").text
                    logger.info("--------------------------------------------------------")

                else:
                    lobpatientcount = driver.find_element_by_xpath(
                        "//*[@id='quality_registry']/div/div[1]/div[3]/div[2]").text
                    logger.info("--------------------------------------------------------")
            except Exception as e:
                lobpatientcount = driver.find_element_by_xpath(
                    "//*[@id='quality_registry']/div/div[1]/div[3]/div[2]").text
                logger.info("--------------------------------------------------------")

            if (lobpatientcount == "0"):
                logger.critical(
                    "Registry  -> " + str(currentLOBName) + " Patient count is 0.Please check.")
                logger.info(
                    "Registry  -> " + str(currentLOBName) + " Patient count is 0.")
                logger.info("--------------------------------------------------------")
            else:
                logger.info(
                    "Registry  ->  " + str(
                        currentLOBName) + " Patient count is : " + str(lobpatientcount))
                logger.info("--------------------------------------------------------")

            # Accordian metric validation started
            try:
                time.sleep(1)
                driver.find_element_by_xpath("//*[@id='metric_scorecard']/div/div[1]/div/span/a[3]").click()
                time.sleep(2)
                driver.find_element_by_xpath("//*[@id='qt-reg-nav-filters']/li[1]/label").click()
                time.sleep(2)
                driver.find_element_by_xpath("//*[@id='qt-apply-search']").click()
                time.sleep(2)
                total_accordion_metric = driver.find_elements_by_xpath("//*[@class='accordion active']")
                print("Total Accordion Metric(s) :  " + str(len(total_accordion_metric)))
                logger.info("Total Accordion Metric(s) :  " + str(len(total_accordion_metric)))

                accordion_metric_list = []
                for i in range(0, len(total_accordion_metric)):

                    print("Accordion Metric Id : " + total_accordion_metric[i].get_attribute('id'))
                    logger.info("Accordion Metric Id : " + total_accordion_metric[i].get_attribute('id'))
                    print(total_accordion_metric[i].get_attribute('id'))

                    # ["382","212","2053","2052","497","85"] -- Corresponding accordion metric id validation have been skipped
                    if (total_accordion_metric[i].get_attribute('id') in ["382", "212", "2053", "2052", "497",
                                                                          "85"]):
                        print("Accordion Metric id have been skipped")

                    else:
                        accordion_metric_list.append(total_accordion_metric[i].get_attribute('id'))
                    # print(accordion_metric_list)
                print("----------------------------------------------------------------------------")
                logger.info("--------------------------------------------------------")
                print("----------------------------------------------------------------------------")
                logger.info("--------------------------------------------------------")

                for i in range(0, len(accordion_metric_list)):
                    parent_num_den = driver.find_element_by_xpath("//*[@id='" + str(
                        accordion_metric_list[i]) + "']/div[1]/div/a/div/div[1]/div[2]/div[2]/span[2]")
                    print("Measure Num/Denom score of the Parent metric id  " + str(
                        accordion_metric_list[i]) + "  :  " + parent_num_den.text)
                    logger.info("Measure Num/Denom score of the Parent metric id  " + str(
                        accordion_metric_list[i]) + "  :  " + parent_num_den.text,
                                )
                    parent_score = parent_num_den.text
                    parent_num_den_extract = re.search('\(([^)]+)', parent_score).group(1)
                    # print(parent_num_den_extract)
                    parent_num_den_val = parent_num_den_extract.replace(",", "")
                    parent_num_den_split = parent_num_den_val.split("/", 1)
                    parent_num_value = parent_num_den_split[0]
                    print("Numerator value of the Parent metric id  " + str(
                        accordion_metric_list[i]) + "  :  " + parent_num_value)
                    logger.info("Numerator value of the Parent metric id  " + str(
                        accordion_metric_list[i]) + "  :  " + parent_num_value,
                                )
                    parent_den_value = parent_num_den_split[1]
                    print("Denominator value of the Parent metric id  " + str(
                        accordion_metric_list[i]) + "  :  " + parent_den_value)
                    logger.info("Denominator value of the Parent metric id  " + str(
                        accordion_metric_list[i]) + "  :  " + parent_den_value,
                                )

                    child_metric = driver.find_elements_by_xpath("//*[@id='" + str(
                        accordion_metric_list[i]) + "']/div[2]//ancestor::div[@class='qt-metric']")
                    print("Total Child Measures of the Parent metric id  " + str(
                        accordion_metric_list[i]) + " :  " + str(len(child_metric)))
                    logger.info("Total Child Measures of the Parent metric id  " + str(
                        accordion_metric_list[i]) + " :  " + str(len(child_metric)),
                                )
                    child_sum_num = 0
                    child_sum_den = 0
                    for j in range(0, len(child_metric)):
                        j = j + 1
                        child_num_den = driver.find_element_by_xpath(
                            "//*[@id='" + str(accordion_metric_list[i]) + "']/div[2]/div[" + str(
                                j) + "]/a/div/div[1]/div[2]/div[2]/span[2]")
                        # print("//*[@id='"+str(accordian_metric_list[i])+"']/div[2]/div["+str(j)+"]/a/div/div[1]/div[2]/div[2]/span[2]")
                        print("Child [" + str(j) + "] Num/Den score : " + child_num_den.text)
                        logger.info("Child [" + str(j) + "] Num/Den score : " + child_num_den.text,
                                    )
                        child_score = child_num_den.text
                        child_num_den_extract = re.search('\(([^)]+)', child_score).group(1)
                        child_num_den_val = child_num_den_extract.replace(",", "")
                        child_num_den_val = child_num_den_extract.replace(",", "")
                        child_num_den_split = child_num_den_val.split("/", 1)
                        child_num_value = child_num_den_split[0]
                        # print(child_num_value)
                        # print(int(child_num_value))
                        child_sum_num = int(child_num_value) + child_sum_num
                        # print(child_sum_num)
                        child_den_value = child_num_den_split[1]
                        # print(child_den_value)
                        child_sum_den = int(child_den_value) + child_sum_den
                        # print(child_sum_den)
                    print("Total Numerator score of the all child metric(s) :  " + str(child_sum_num))
                    logger.info("Total Numerator score of the all child metric(s) :  " + str(child_sum_num),
                                )
                    print("Total Denominator score of the all child metric(s) :  " + str(child_sum_den))
                    logger.info("Total Denominator score of the all child metric(s) :  " + str(child_sum_den),
                                )
                    if (int(parent_num_value) == child_sum_num and int(parent_den_value) == child_sum_den):
                        print("Sum of child score is matching with parent score")
                        logger.info("Sum of child score is matching with parent score",
                                    )
                    else:
                        print("Score didn't matched")
                        logger.info("###### Score didn't matched ######")
                        logger.critical("LOB Selected  :: " + currentLOBName)
                        logger.critical("---------------------------------------")
                        logger.critical("Parent metric id  :   " + str(accordion_metric_list[i]))
                        logger.critical("Score didn't matched for the corresponding Parent metric id",
                                        )
                        logger.critical("--------------------------------------------------------",
                                        )
                    print("----------------------------------------------------------------------------")
                    logger.info("--------------------------------------------------------")



            except Exception as e:
                print(e)

            time.sleep(1)
            LOBdropdownelement = driver.find_element_by_xpath("//*[@id='qt-filter-label']")
            LOBdropdownelement.click()
            time.sleep(1)
            LOBname = LOBdropdownelement.find_element_by_xpath("//*[@id='filter-lob']")
            LOBnamelist = LOBname.find_elements_by_tag_name("li")

     except Exception as e:
         logger.critical(
             "Registry  -> Accordion Measure validation have been suspended due to error!!Please check.")
'''


def SupportpageAccordionValidation(driver, workbook, logger, run_from):
    try:
        workbook.create_sheet('Accordian Validation')
        ws = workbook['Accordian Validation']

        ws.append(['ID', 'LoB Name', 'Metric ID', 'Status', 'Comments'])
        header_font = Font(color='FFFFFF', bold=False, size=12)
        header_cell_color = PatternFill('solid', fgColor='030303')
        ws['A1'].font = header_font
        ws['A1'].fill = header_cell_color
        ws['B1'].font = header_font
        ws['B1'].fill = header_cell_color
        ws['C1'].font = header_font
        ws['C1'].fill = header_cell_color
        ws['D1'].font = header_font
        ws['D1'].fill = header_cell_color
        ws['E1'].font = header_font
        ws['E1'].fill = header_cell_color
        ws.name = "Arial"
        test_case_id = 1
        loader = WebDriverWait(driver, 300)
        loader.until(EC.invisibility_of_element_located((By.XPATH, "//div/div[contains(@class,'ajax_preloader')]")))
        LOBdropdownelement = driver.find_element_by_xpath("//*[@id='qt-filter-label']")
        LOBdropdownelement.click()
        default_quarter = driver.find_element_by_xpath(
            "//*[@id='filter-quarter']//following-sibling::li[@class=' highlight ']/span/a")
        print(default_quarter.text)
        logger.info("Default Quarter selected as  : " + default_quarter.text)
        logger.info("--------------------------------------------------------")
        Flag = locator.select_measurement_year_flag_support
        # Navigation with customize Measurement Year
        # LOBquarterlist = []
        print(Flag)
        LOBquarter = LOBdropdownelement.find_elements_by_xpath("//*[@id='filter-quarter']/li")
        if (Flag == "True"):
            for i in range(0, len(LOBquarter)):
                # LOBquarterlist.append(LOBquarter[i].text)
                if LOBquarter[i].text == locator.MeasurementYear_Support or LOBquarter[i].text == locator.MeasurementYearQuarter_Support:
                    logger.info("Current Quarter selected as  : " + LOBquarter[i].text)
                    logger.info("--------------------------------------------------------")
                    LOBquarter[i].click()
                    break
                LOBquarter = LOBdropdownelement.find_elements_by_xpath("//*[@id='filter-quarter']/li")

        time.sleep(1)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='filter-lob']")))
        LOBname = LOBdropdownelement.find_element_by_xpath("//*[@id='filter-lob']")
        LOBnamelist = LOBname.find_elements_by_tag_name("li")
        print(*LOBnamelist)
        Payername = LOBdropdownelement.find_elements_by_xpath("//*[@id='filter-payer']")
        # LOBdropdownelement.click()
        time.sleep(1)
        for j in range(0, len(LOBnamelist)):
            # LOBdropdownelement.click()
            time.sleep(1)
            print(LOBnamelist[j].text)
            print("--------------------------------")
            LOBnamelist[j].click()
            currentLOBName = LOBnamelist[j].text
            logger.info("LOB Selected  :: " + currentLOBName)
            logger.info("---------------------------------------")
            driver.find_element_by_xpath("//*[@id='reg-filter-apply']").click()
            time.sleep(2)
            loader = WebDriverWait(driver, 300)
            loader.until(
                EC.invisibility_of_element_located((By.XPATH, "//div/div[contains(@class,'ajax_preloader')]")))
            # logger.captureScreenshot(driver, currentLOBName, screenshot_path)
            # Checking Patient count in Lob. Raise error if it is 0
            Lob_type = ["ALL", "Medicare", "Medicare ACO"]
            try:
                if currentLOBName in Lob_type:
                    lobpatientcount = driver.find_element_by_xpath(
                        "//*[@id='quality_registry']/div/div[1]/div[4]/div[2]").text
                    logger.info("--------------------------------------------------------")

                else:
                    lobpatientcount = driver.find_element_by_xpath(
                        "//*[@id='quality_registry']/div/div[1]/div[3]/div[2]").text
                    logger.info("--------------------------------------------------------")
            except Exception as e:
                lobpatientcount = driver.find_element_by_xpath(
                    "//*[@id='quality_registry']/div/div[1]/div[3]/div[2]").text
                logger.info("--------------------------------------------------------")

            if (lobpatientcount == "0"):
                logger.critical(
                    "Registry  -> " + str(currentLOBName) + " Patient count is 0.Please check.")
                logger.info(
                    "Registry  -> " + str(currentLOBName) + " Patient count is 0.")
                logger.info("--------------------------------------------------------")
            else:
                logger.info(
                    "Registry  ->  " + str(
                        currentLOBName) + " Patient count is : " + str(lobpatientcount))
                logger.info("--------------------------------------------------------")
                try:

                    time.sleep(1)
                    driver.find_element_by_xpath("//*[@id='metric_scorecard']/div/div[1]/div/span/a[3]").click()
                    time.sleep(2)
                    driver.find_element_by_xpath("//*[@id='qt-reg-nav-filters']/li[1]/label").click()
                    time.sleep(2)
                    driver.find_element_by_xpath("//*[@id='qt-apply-search']").click()
                    time.sleep(2)
                    total_accordion_metric = driver.find_elements_by_xpath("//*[@class='accordion active']")
                    print("Total Accordion Metric(s) :  " + str(len(total_accordion_metric)))
                    logger.info("Total Accordion Metric(s) :  " + str(len(total_accordion_metric)))

                    accordion_metric_list = []
                    for i in range(0, len(total_accordion_metric)):

                        print("Accordion Metric Id : " + total_accordion_metric[i].get_attribute('id'))
                        logger.info("Accordion Metric Id : " + total_accordion_metric[i].get_attribute('id'))
                        print(total_accordion_metric[i].get_attribute('id'))

                        # ["382","212","2053","2052","497","85"] -- Corresponding accordion metric id validation have been skipped
                        if (total_accordion_metric[i].get_attribute('id') in ["382", "212", "2053", "2052", "497",
                                                                              "85"]):
                            print("Accordion Metric id have been skipped")

                        else:
                            accordion_metric_list.append(total_accordion_metric[i].get_attribute('id'))
                        # print(accordion_metric_list)
                    print("----------------------------------------------------------------------------")
                    logger.info("--------------------------------------------------------")
                    print("----------------------------------------------------------------------------")
                    logger.info("--------------------------------------------------------")

                    for i in range(0, len(accordion_metric_list)):
                        parent_num_den = driver.find_element_by_xpath("//*[@id='" + str(
                            accordion_metric_list[i]) + "']/div[1]/div/a/div/div[1]/div[2]/div[2]/span[2]")
                        print("Measure Num/Denom score of the Parent metric id  " + str(
                            accordion_metric_list[i]) + "  :  " + parent_num_den.text)
                        logger.info("Measure Num/Denom score of the Parent metric id  " + str(
                            accordion_metric_list[i]) + "  :  " + parent_num_den.text, )
                        parent_score = parent_num_den.text
                        parent_num_den_extract = re.search('\(([^)]+)', parent_score).group(1)
                        # print(parent_num_den_extract)
                        parent_num_den_val = parent_num_den_extract.replace(",", "")
                        parent_num_den_split = parent_num_den_val.split("/", 1)
                        parent_num_value = parent_num_den_split[0]
                        print("Numerator value of the Parent metric id  " + str(
                            accordion_metric_list[i]) + "  :  " + parent_num_value)
                        logger.info("Numerator value of the Parent metric id  " + str(
                            accordion_metric_list[i]) + "  :  " + parent_num_value,
                                    )
                        parent_den_value = parent_num_den_split[1]
                        print("Denominator value of the Parent metric id  " + str(
                            accordion_metric_list[i]) + "  :  " + parent_den_value)
                        logger.info("Denominator value of the Parent metric id  " + str(
                            accordion_metric_list[i]) + "  :  " + parent_den_value,
                                    )

                        child_metric = driver.find_elements_by_xpath("//*[@id='" + str(
                            accordion_metric_list[i]) + "']/div[2]//ancestor::div[@class='qt-metric']")
                        print("Total Child Measures of the Parent metric id  " + str(
                            accordion_metric_list[i]) + " :  " + str(len(child_metric)))
                        logger.info("Total Child Measures of the Parent metric id  " + str(
                            accordion_metric_list[i]) + " :  " + str(len(child_metric)),
                                    )
                        child_sum_num = 0
                        child_sum_den = 0
                        for j in range(0, len(child_metric)):
                            j = j + 1
                            child_num_den = driver.find_element_by_xpath(
                                "//*[@id='" + str(accordion_metric_list[i]) + "']/div[2]/div[" + str(
                                    j) + "]/a/div/div[1]/div[2]/div[2]/span[2]")
                            # print("//*[@id='"+str(accordian_metric_list[i])+"']/div[2]/div["+str(j)+"]/a/div/div[1]/div[2]/div[2]/span[2]")
                            print("Child [" + str(j) + "] Num/Den score : " + child_num_den.text)
                            logger.info("Child [" + str(j) + "] Num/Den score : " + child_num_den.text,
                                        )
                            child_score = child_num_den.text
                            child_num_den_extract = re.search('\(([^)]+)', child_score).group(1)
                            child_num_den_val = child_num_den_extract.replace(",", "")
                            child_num_den_val = child_num_den_extract.replace(",", "")
                            child_num_den_split = child_num_den_val.split("/", 1)
                            child_num_value = child_num_den_split[0]
                            # print(child_num_value)
                            # print(int(child_num_value))
                            child_sum_num = int(child_num_value) + child_sum_num
                            # print(child_sum_num)
                            child_den_value = child_num_den_split[1]
                            # print(child_den_value)
                            child_sum_den = int(child_den_value) + child_sum_den
                            # print(child_sum_den)
                        print("Total Numerator score of the all child metric(s) :  " + str(child_sum_num))
                        logger.info("Total Numerator score of the all child metric(s) :  " + str(child_sum_num),
                                    )
                        print("Total Denominator score of the all child metric(s) :  " + str(child_sum_den))
                        logger.info("Total Denominator score of the all child metric(s) :  " + str(child_sum_den),
                                    )
                        if (int(parent_num_value) == child_sum_num and int(parent_den_value) == child_sum_den):
                            print("Sum of child score is matching with parent score")
                            logger.info("Sum of child score is matching with parent score", )
                            ws.append([test_case_id, currentLOBName, accordion_metric_list[i], "Passed",
                                       "Sum of child score is matching with parent score. Parent Score: " + parent_num_den.text + " ,Child score sum: (" + str(
                                           child_sum_num) + "/" + str(child_sum_den) + ")"])
                            test_case_id += 1
                        else:
                            print("Score didn't matched")
                            logger.info("###### Score didn't matched ######")
                            logger.critical("LOB Selected  :: " + currentLOBName)
                            logger.critical("---------------------------------------")
                            logger.critical("Parent metric id  :   " + str(accordion_metric_list[i]))
                            logger.critical("Score didn't matched for the corresponding Parent metric id",
                                            )
                            logger.critical("--------------------------------------------------------",
                                            )
                            ws.append([test_case_id, currentLOBName, accordion_metric_list[i], "Failed",
                                       "Sum of child score is not matching with parent score. Parent Score: " + parent_num_den.text + " ,Child score sum: (" + str(
                                           child_sum_num) + "/" + str(child_sum_den) + ")"])
                            test_case_id += 1

                        print("----------------------------------------------------------------------------")
                        logger.info("--------------------------------------------------------")




                except Exception as e:
                    print(e)

                time.sleep(1)
                LOBdropdownelement = driver.find_element_by_xpath("//*[@id='qt-filter-label']")
                LOBdropdownelement.click()
                time.sleep(1)
                LOBname = LOBdropdownelement.find_element_by_xpath("//*[@id='filter-lob']")
                LOBnamelist = LOBname.find_elements_by_tag_name("li")

    except Exception as e:
        traceback.print_exc()
        logger.critical(
            "Registry  -> Accordion Measure validation have been suspended due to error!!Please check.")

    rows = ws.max_row
    cols = ws.max_column
    for i in range(2, rows + 1):
        for j in range(3, cols + 1):
            if ws.cell(i, j).value == 'Passed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='0FC404')
            elif ws.cell(i, j).value == 'Failed':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FC0E03')
            elif ws.cell(i, j).value == 'Data table is empty':
                ws.cell(i, j).fill = PatternFill('solid', fgColor='FCC0BB')
