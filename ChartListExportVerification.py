import base64
import glob
import os
import shutil
import traceback

import variablestorage as locator
import setups
import filter_handler_cozeva as filterValidator

import pyautogui
import re
import time
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from dateutil import parser
from openpyxl.styles import PatternFill,Font
### Write in such a way that you can always edit the page you will compare for time load
from selenium import webdriver
from selenium.common.exceptions import ElementNotInteractableException, ElementClickInterceptedException, NoSuchElementException, \
    TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import ChartListExport
from openpyxl import Workbook

#main_chart_list_export(customer_id, file_path_supplemental_data, file_path_hcc_chart_list, file_path_awv_chart_list, file_path_report)
# Design -------------
def make_directory(path1):
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


#setup driver
download_dir_original="C:\\Users\\ssrivastava\\Downloads"
download_dir_original = locator.download_dir
# Get the current date and time
now = datetime.now()

# Format the date and time as YYYY-MM-DD-HH-MM-SS
timestamp = now.strftime("%Y-%m-%d-%H-%M-%S")

# Create a directory with the timestamp as the name
directory_name_download = f"Export_Download_{timestamp}"

os.makedirs(os.path.join(download_dir_original,directory_name_download))
download_dir=os.path.join(download_dir_original,directory_name_download)

# report_place="C:\\VerificationReports"
# directory_name_reports=f"ChartList_Export_Reports_{timestamp}"
# os.makedirs(os.path.join(report_place,directory_name_reports))
# report=os.path.join(report_place,directory_name_reports)
environment_name="PROD"
baseurl="https://www.cozeva.com"

def wait_to_load(driver,timeout):
    loader_element_class='ajax_preloader'
    WebDriverWait(driver,timeout).until(EC.invisibility_of_element_located((By.CLASS_NAME, loader_element_class)))

def action_click(driver,element):
    try:
        element.click()
    except (ElementNotInteractableException, ElementClickInterceptedException):
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        driver.execute_script("arguments[0].click();", element)

def delete_folder(folder_path):
    # Check if the folder exists
    if os.path.exists(folder_path):
        # Delete the folder and all its contents
        shutil.rmtree(folder_path)
        print("Folder deleted successfully.")
    else:
        print("Folder does not exist.")

def download_and_copy_files(customer_id,chart_type):



    # After ChartListExport Customer Folder is freshly created everytime
    customer_export_dir = f"C:\\ChartListExports\\{customer_id}"
    os.makedirs(customer_export_dir, exist_ok=True)
    print(customer_export_dir, " Created")

    # Check the download folder for Supplemental data export

    # Get a list of all files in the downloads folder
    files = glob.glob(os.path.join(download_dir, '*'))

    # Sort files by creation time (latest first)
    files.sort(key=os.path.getmtime, reverse=True)

    # Flag to track if the file is found
    file_found = False

    # Iterate through the files # Get the latest downloaded file based on modification timeSupplemental Data
    for file_path in files:
        if chart_type in file_path:
            print(f"File containing {chart_type} found: {file_path}")
            latest_file = file_path
            file_found = True
            break

    if not file_found:
        print("No file containing 'Supplemental Data' found.")

    #name of the file
    downloaded_file_name = os.path.basename(latest_file)

    # Path to the destination file
    destination_folder = customer_export_dir

    # Copy the file to the destination directory
    shutil.copy(latest_file, destination_folder)
    print(f"Latest downloaded file '{latest_file}' moved to '{destination_folder}' successfully.")


    print(f"Customer {customer_id} file available at {os.path.join(destination_folder,downloaded_file_name)}")
    return os.path.join(destination_folder,downloaded_file_name)




def check_chartlist_export(driver,customer_id,wb, report_folder, workbook_title):
    hamburger_icon="//i[text()='menu']"
    supplemental_data_link_xpath="//li[@class='chart_chase_list_type' and @data-list-type='1']//a"
    hcc_data_link_xpath="//li[@class='chart_chase_list_type' and @data-list-type='2']//a"
    awv_data_link_xpath="//li[@class='chart_chase_list_type' and @data-list-type='3']//a"
    column_header_xpath="//th[2]"
    filter_list_xpath = "//i[text()=\"filter_list\"]"
    new_creation_date_filter_from_xpath = "//input[@name='chart_chase_uploaded_from']"
    new_creation_date_filter_to_xpath = "//input[@name='chart_chase_uploaded_to']"
    apply_xpath = "//a[@class='rfloat green-text datatable_apply' and text()='Apply']"
    footer_xpath="//div[@class='dataTables_info']"
    export_icon_xpath="//a[@data-tooltip=\"Export\"]"
    export_list_xpath="//a[text()='Export all to CSV ']"
    # export_option_xpath=

    #make a directory of customer_id in  C://ChartListExports//customer_id
    # Path to the directory


    # # Check if the directory exists
    # if os.path.exists(directory_path):
    #     print(f"Directory '{directory_path}' exists")
    # else:
    #     # Create the directory
    #     os.makedirs(directory_path)
    #     print(f"Created Directory '{directory_path}'.")

    #Open Registry page for customer
    customer_list_url = []
    sm_customer_id = customer_id
    session_var = 'app_id=registries&custId=' + str(sm_customer_id) + '&payerId=' + str(
        sm_customer_id) + '&orgId=' + str(sm_customer_id) + '&vgpId=' + str(sm_customer_id) + '&vpId=' + str(
        sm_customer_id)
    encoded_string = base64.b64encode(session_var.encode('utf-8'))
    customer_list_url.append(encoded_string)
    for idx, val in enumerate(customer_list_url):
        url = (baseurl+"/registries?session=" + val.decode('utf-8'))

    customer_id_file_dict={}
    customer_id_file_dict[customer_id]=[]
    idx=1
    while(idx<=3):
        if (idx == 1):
            xpath_to_click = supplemental_data_link_xpath
            chart_type = "Supplemental Data"
        if (idx == 2):
            xpath_to_click = hcc_data_link_xpath
            chart_type = "HCC Chart"
        if (idx == 3):
            xpath_to_click = awv_data_link_xpath
            chart_type = "AWV Chart"
        driver.get(url)
        # open supplemental data chart list
        wait_to_load(driver, 300)
        action_click(driver, driver.find_element(By.XPATH, hamburger_icon))
        try:
            action_click(driver, driver.find_element(By.XPATH, xpath_to_click))
            wait_to_load(driver, 300)
            timeout_for_column_headers = 20
            WebDriverWait(driver, timeout_for_column_headers).until(
                EC.visibility_of_element_located((By.XPATH, column_header_xpath)))
            list_base_url = driver.current_url
            try:
                if chart_type == "Supplemental Data" and (customer_id == "1300" or customer_id == "3000"):
                    filterValidator.supplemental_data_list(driver, customer_id, wb, report_folder, workbook_title)
                elif chart_type == "HCC Chart" and (customer_id == "1300" or customer_id == "3000"):
                    filterValidator.hcc_chart_list(driver, customer_id, wb, report_folder, workbook_title)
                elif chart_type == "AWV Chart" and (customer_id == "1300" or customer_id == "3000"):
                    filterValidator.awv_chart_list(driver, customer_id, wb, report_folder, workbook_title)
                driver.get(list_base_url)
                wait_to_load(driver, 300)
                timeout_for_column_headers = 20
                WebDriverWait(driver, timeout_for_column_headers).until(
                    EC.visibility_of_element_located((By.XPATH, column_header_xpath)))
            except Exception as e:
                traceback.print_exc()
                driver.get(list_base_url)
                wait_to_load(driver, 300)
                timeout_for_column_headers = 20
                WebDriverWait(driver, timeout_for_column_headers).until(
                    EC.visibility_of_element_located((By.XPATH, column_header_xpath)))
        except NoSuchElementException:
            print(f"Error : Unable to find chart_type {chart_type}")
            stored_file_path = '9999'
            print(f"Customer {customer_id} file available at {stored_file_path}")
            customer_id_file_dict[customer_id].append(str(stored_file_path))
            if(idx==2):
                customer_id_file_dict[customer_id].append(str(stored_file_path))
            break
        time_delta = 3
        while (1):
            # wait for page to load
            wait_to_load(driver, 300)
            timeout_for_column_headers = 20
            WebDriverWait(driver, timeout_for_column_headers).until(
                EC.visibility_of_element_located((By.XPATH, column_header_xpath)))



            # Set date filter
            # Get the current date
            current_date = datetime.now()-relativedelta(months=2)
            formatted_date_to = current_date.strftime("%m/%d/%Y")
            date_from = current_date - timedelta(days=time_delta)
            formatted_date_from = date_from.strftime("%m/%d/%Y")

            WebDriverWait(driver, timeout_for_column_headers).until(
                EC.element_to_be_clickable((By.XPATH, filter_list_xpath)))
            # first apply date filter
            action_click(driver, driver.find_element(By.XPATH, filter_list_xpath))
            time.sleep(2)
            driver.find_element(By.XPATH, new_creation_date_filter_from_xpath).clear()
            time.sleep(1)
            driver.find_element(By.XPATH, new_creation_date_filter_from_xpath).send_keys(formatted_date_from)

            driver.find_element(By.XPATH, new_creation_date_filter_to_xpath).clear()
            time.sleep(1)
            driver.find_element(By.XPATH, new_creation_date_filter_to_xpath).send_keys(formatted_date_to)

            try:
                action_click(driver, driver.find_element(By.XPATH, apply_xpath))
                print("Clicked on applied")
            except Exception as e:
                print("Unable to click on Apply")

            wait_to_load(driver, 300)
            time.sleep(2)
            WebDriverWait(driver, timeout_for_column_headers).until(
                EC.visibility_of_element_located((By.XPATH, footer_xpath)))

            # check number of entries
            time.sleep(5)
            footer_text = driver.find_element(By.XPATH, footer_xpath).get_attribute("innerHTML")
            # Find the index of "of" and "entries"
            index_of_of = footer_text.find("of")
            index_of_entries = footer_text.find("entries")

            # Extract the number between "of" and "entries"
            number_between_of_and_entries = int(footer_text[index_of_of + 3:index_of_entries].strip().replace(",", ""))

            print("No of entries ", number_between_of_and_entries)
            if (number_between_of_and_entries >=300 and number_between_of_and_entries < 2500):
                print(f"Optimal entries present for {time_delta} days before current date ")
                break
            if (number_between_of_and_entries < 300):
                time_delta = time_delta + 5
                if(chart_type=="HCC Chart" or (chart_type=="AWV Chart" and number_between_of_and_entries>0) ):
                    print(f"Optimal entries present for {time_delta} days before current date ")
                    break

            if (number_between_of_and_entries > 2500):
                time_delta = time_delta - 1

            # download file to C://Downloads//ChartList//CustomerID//__.csv
        download_successful = False
        try:
            action_click(driver, driver.find_element(By.XPATH, export_icon_xpath))
            action_click(driver, driver.find_element(By.XPATH, export_list_xpath))
            download_successful = True
            print("Downloaded File")
            time.sleep(30)


        except Exception as e:
            print("Error occurred in downloading file " + str(e))

        if (download_successful):
            # Creates the folder ChartListExport if it is not there

            stored_file_path = download_and_copy_files(customer_id, chart_type)
            print(f"Customer {customer_id} file available at {stored_file_path}")
            customer_id_file_dict[customer_id].append(stored_file_path)


    #fix here for optum
        idx=idx+1
    return customer_id_file_dict




#login to baseurl

# Valid values STAGE , PROD , CERT

# logout_url = baseurl+"/user/logout"
# login_url = baseurl+"/user/login"
#
#
# driver=setup()
# driver.get(logout_url)
# driver.get(login_url)
# creds = get_credentials(environment_name)
#
# #login
# uname = driver.find_element(By.ID,"edit-name")
# pwd = driver.find_element(By.ID,"edit-pass")
# uname.send_keys(creds[0])
# pwd.send_keys(creds[1])
# driver.find_element(By.ID,"edit-submit").click()
# # reason for login
# WebDriverWait(driver, 120).until(EC.presence_of_element_located((By.XPATH, "//textarea[@id=\"reason_textbox\"]")))
# actions = ActionChains(driver)
# reason = driver.find_element(By.XPATH,"//textarea[@id=\"reason_textbox\"]")
# actions.click(reason)
# actions.send_keys_to_element(reason, "https://redmine2.cozeva.com/issues/7662 ")
# actions.perform()
# driver.find_element(By.ID,"edit-submit").click()
# print("Logged in")

#"3000","1300","200","4600","6800","6700","1000","3300","1850

#customer_ids = ["3000","1300","200","4600","6800","6700","1000","3300","1850"]
customer_ids = ["3000", "1300"]
customer_ids.sort()

start_time = time.time()
print("Execution started")
export_dir = f"C:\\VerificationReports\\ChartListExports"
delete_folder(export_dir)
os.mkdir(export_dir)
report='C:/VerificationReports/Chart List Reports/'
filter_workbook, report_folder, workbook_title = filterValidator.create_report_folder_and_file(["3000", "1300"])
filter_workbook.save(report_folder + "\\" + workbook_title)
isdir = os.path.isdir(report)
if not isdir:
    os.mkdir(report)
driver = setups.driver_setup_with_download(download_dir)
setups.login_to_cozeva("1500")


for customer_id in customer_ids:
    #print(check_chartlist_export(driver, customer_id)) ['3000', 'C:\\ChartListExports\\3000\\Supplemental Data 2024-04-18 (1).csv', '0', '0', 'C:\\VerificationReports\\']
    #parameters=check_chartlist_export(driver, customer_id)

    input_dict=check_chartlist_export(driver, customer_id, filter_workbook, report_folder, workbook_title)
    ChartListExport.main_chart_list_export(int(customer_id),input_dict[customer_id][0],input_dict[customer_id][1],input_dict[customer_id][2],report)

driver.quit()
exec(open("view_chartlist_report.py").read(), globals())






# Hard code and run
# customer_id=1300
# supp_data_path='C:/Users/ssrivastava/Downloads/Export_Download_2024-05-16-23-30-49/Supplemental Data 2024-05-16.csv'
# hcc_data_path='C:/Users/ssrivastava/Downloads/Export_Download_2024-05-16-23-30-49/HCC Chart List 2024-05-16.csv'
# awv_data_path='C:/Users/ssrivastava/Downloads/Export_Download_2024-05-16-23-30-49/AWV Chart List 2024-05-16.csv'
# #awv_data_path=''



#ChartListExport.main_chart_list_export(customer_id,supp_data_path,hcc_data_path,awv_data_path,report)



# end_time = time.time()
# execution_time = end_time - start_time
# print(f"Execution time: {execution_time} seconds")
