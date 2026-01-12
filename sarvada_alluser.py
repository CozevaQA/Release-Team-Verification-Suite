import time
from selenium.common.exceptions import ElementNotInteractableException, ElementClickInterceptedException, NoSuchElementException, \
    TimeoutException, UnexpectedAlertPresentException, StaleElementReferenceException
from selenium.webdriver.common.by import By
import pandas as pd
import os
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
from colorama import Fore, Style
from datetime import datetime
#Reporting


def create_html_report(user_id, all_dictionary, other_dictionary, report_directory):
    # Ensure the report directory exists
    os.makedirs(report_directory, exist_ok=True)

    # Get all unique keys from both dictionaries
    all_keys = set(all_dictionary.keys()).union(set(other_dictionary.keys()))

    # Start building the HTML content
    html_content = f"""
    <html>
    <head>
        <title>Comparison Report - {user_id}</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 20px; }}
            h2 {{ text-align: center; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ border: 1px solid black; padding: 8px; text-align: left; }}
            th {{ background-color: #f2f2f2; }}
            .pass {{ background-color: #c6efce; color: #006100; }}
            .fail {{ background-color: #ffc7ce; color: #9c0006; }}
        </style>
    </head>
    <body>
        <h2>Comparison Report for User: {user_id}</h2>
        <table>
            <tr>
                <th>Key</th>
                <th>All LOB Value</th>
                <th>Other LOB Value</th>
                <th>Status</th>
            </tr>
    """

    # Compare values and add rows to the HTML table
    for key in sorted(all_keys):
        val1 = all_dictionary.get(key, "N/A")
        val2 = other_dictionary.get(key, "N/A")
        status = "Pass" if val1 == val2 else "Fail"
        status_class = "pass" if status == "Pass" else "fail"

        html_content += f"""
            <tr>
                <td>{key}</td>
                <td>{val1}</td>
                <td>{val2}</td>
                <td class="{status_class}">{status}</td>
            </tr>
        """

    # Close the HTML content
    html_content += """
        </table>
    </body>
    </html>
    """

    # Define the report file path
    timestamp=datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    report_file = os.path.join(report_directory, f"comparison_report_{user_id}_{timestamp}.html")

    # Write the HTML content to a file
    with open(report_file, "w", encoding="utf-8") as file:
        file.write(html_content)

    print(f"Report generated successfully: {report_file}")
    absolute_path = os.path.abspath(report_file).replace("\\", "/")
    report_url = f'file:///{absolute_path}'
    print(f"Report link  successfully: {report_url}")
    return report_url


def flatten_dict(d, parent_key=""):
    """Recursively flattens a nested dictionary."""
    items = []
    for k, v in d.items():
        new_key = f"{parent_key}.{k}" if parent_key else k  # Maintain hierarchy
        if isinstance(v, dict):
            items.extend(flatten_dict(v, new_key).items())  # Recurse for nested dicts
        else:
            items.append((new_key, v))
    return dict(items)

# LOB and Context
#create_html_report("User456", flatten_dict(new_dict1), flatten_dict(new_dict2), "./reports")





def append_comparison_report(all_dict, other_dict, file_path):
    """
    Compares two dictionaries and appends an HTML report to the given file.

    :param all_dict: Dictionary containing data for 'ALL'
    :param other_dict: Dictionary containing data for 'OTHER'
    :param file_path: File path to append the HTML report
    """

    # Prepare the report data
    report_data = []

    # Combine keys from both dictionaries
    all_keys = sorted(set(all_dict.keys()).union(set(other_dict.keys())))

    # Process each key
    for key in all_keys:
        # Split the key into LOB, Plan Type, and Metric Abbreviation
        parts = key.split('&')
        #print(parts)
        lob = parts[0]
        plan_type = parts[1]
        metric_abb = parts[2]

        # Get values from both dictionaries or default to [None, None]
        all_values = all_dict.get(key, [0.0, 0.0])
        other_values = other_dict.get(key, [0.0, 0.0])

        # Determine the status and apply color coding
        if all_values == other_values:
            status = '<span style="color: green; font-weight: bold;">Pass</span>'
        else:
            status = '<span style="color: red; font-weight: bold;">Fail</span>'

        # Append the row to the report
        # print(all_values)
        if(len(all_values)==1):
            all_values.append(all_values[0])

        if (len(other_values) == 1):
            other_values.append(other_values[0])


        report_data.append([lob, plan_type, metric_abb,
                            all_values[0], all_values[1],
                            other_values[0], other_values[1],
                            status])

    # Create a DataFrame for the report
    df = pd.DataFrame(report_data, columns=['LOB', 'PLAN-TYPE', 'METRIC ABBREVIATION',
                                            'NUMERATOR FROM ALL', 'DENOMINATOR FROM ALL',
                                            'NUMERATOR FROM OTHER', 'DENOMINATOR FROM OTHER',
                                            'STATUS'])

    # Convert DataFrame to HTML with styling
    html_report = df.to_html(index=False, escape=False)

    # Add custom CSS for styling
    styled_html = f"""
        <style>
            table {{
                width: 100%;
                border-collapse: collapse;
                text-align: center;
            }}
            th {{
                background-color: #ADD8E6; /* Light blue header */
                color: black;
                padding: 8px;
            }}
            td {{
                text-align: center;
                padding: 8px;
            }}
            table, th, td {{
                border: 1px solid #ddd;
            }}
        </style>
        {html_report}
        <br><hr><br>
        """

    # Append HTML to the file

    # Append HTML to the file
    with open(file_path, "a+") as file:
        file.write(html_report)
        file.write("<br><hr><br>")  # Add a separator between reports

    print(f"HTML report appended to {file_path}")


#Reporting ENDS


def action_click(driver,element):
    try:
        element.click()
    except (ElementNotInteractableException, ElementClickInterceptedException,StaleElementReferenceException):
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        driver.execute_script("arguments[0].click();", element)

def ajax_preloader_wait(driver):
    time.sleep(1)
    #WebDriverWait(driver, 300).until(
    #    EC.invisibility_of_element((By.XPATH, "//div/div[contains(@class,'ajax_preloader')]")))
    WebDriverWait(driver, 300).until(
        EC.invisibility_of_element((By.CLASS_NAME, "ajax_preloader")))
    #time.sleep(1)
    if len(driver.find_elements(By.CLASS_NAME,"ajax_preloader")) != 0:
        WebDriverWait(driver, 300).until(
            EC.invisibility_of_element((By.CLASS_NAME, "ajax_preloader")))

    WebDriverWait(driver, 300).until(
        EC.invisibility_of_element((By.CLASS_NAME, "drupal_message_text")))

    time.sleep(1)

def extract_plan_type_from_url(url):
    value="ALL"

    match = re.search(r"&plan_type=([^&]+)", url)
    if match:
        value = match.group(1)
        print(value)
    return value

def extract_numerator_denominator(text):
    # Split by '/' and remove commas
    num_den = [float(num.replace(",", "")) for num in text.split("/")]
   # Output: [1132, 11182]
    return num_den

def extract_metric_lob_info(driver,context_name,year):
     #info -> Context_name | YEAR | Plan_type |  LOB | Measure_Abbr | Numerator | Denominator
    #Steps : Click on green pill : Wait for loader : Extract info

     metric_xpath = "//div[@class=\"qt-metric\"]"
    # Plan Type from URL
    # LOB from green pill
    # Measure_abbreviation f"(//div[@class="qt-metric"])[{}]//child::span[@class="met-abbr"]"
    # Num / den score (//div[@class="qt-metric"])[15]//child::span[@class="num-den"]
     green_pill_xpath = "//a[@id='qt-filter-label']//span[2]"
     year_xpath = f"//ul[@id=\"filter-quarter\"]//a[text()={year}]"
     lob_elements_xpaths = "//ul[@id=\"filter-lob\"]//a"
     count_xpaths = "//ul[@id=\"filter-payer\"]//a[text()=\"ALL\"]//following-sibling::a"
     apply_xpath="//a[text()=\"Apply\"]"
     # click on green pill
     action_click(driver, driver.find_element(By.XPATH, green_pill_xpath))
     # select year
     action_click(driver, driver.find_element(By.XPATH, year_xpath))

     lob_elements = driver.find_elements(By.XPATH, lob_elements_xpaths)
     metric_lob_dict={}
     for lob_element in lob_elements:
         # click on green pill
         action_click(driver, driver.find_element(By.XPATH, green_pill_xpath))
         # select year
         action_click(driver, driver.find_element(By.XPATH, year_xpath))
         time.sleep(2)
         action_click(driver, lob_element)
         lob=lob_element.get_attribute('innerHTML')
         # if(lob=="Medicaid"):
         #     lob="Medi-Cal"
         action_click(driver,driver.find_element(By.XPATH,apply_xpath))
         time.sleep(15)
         plan_type=extract_plan_type_from_url(driver.current_url)
         # print(driver.current_url)
         # print(plan_type)
         metric_number=len(driver.find_elements(By.XPATH,metric_xpath))+1
         num_den=[]
         print("Number of metrics ",metric_number)
         # Find the element with class "accordion li-metric"
         elements = driver.find_elements(By.XPATH, "//i[text()=\"keyboard_arrow_down\"]")
         # Replace the class with "accordion li-metric active"
         for i in range(1,len(elements)+1):
             element_xpath=f"(//i[text()=\"keyboard_arrow_down\"])[{i}]"
             action_click(driver,driver.find_element(By.XPATH,element_xpath))
             time.sleep(2)
         for i in range(1,metric_number):
             metric_abb_xpath=f"(//div[@class=\"qt-metric\"])[{i}]//child::span[@class=\"met-abbr\"]"
             metric_abb=driver.find_element(By.XPATH,metric_abb_xpath).text
             key=lob+"&"+plan_type+"&"+metric_abb
             num_den_xpath=f"(//div[@class=\"qt-metric\"])[{i}]//child::span[@class=\"num-den\"]"
             num_den=extract_numerator_denominator(driver.find_element(By.XPATH,num_den_xpath).get_attribute('innerHTML'))
             metric_lob_dict[key]=num_den
     print(context_name+" ",metric_lob_dict)
     return metric_lob_dict


def extract_green_pill_info(driver,context_name,year):
    green_pill_xpath="//a[@id='qt-filter-label']//span[2]"
    year_xpath=f"//ul[@id=\"filter-quarter\"]//a[text()={year}]"
    lob_elements_xpaths="//ul[@id=\"filter-lob\"]//a"
    count_xpaths="//ul[@id=\"filter-payer\"]//a[text()=\"ALL\"]//following-sibling::a"

    #click on green pill
    action_click(driver,driver.find_element(By.XPATH,green_pill_xpath))
    #select year
    action_click(driver,driver.find_element(By.XPATH,year_xpath))

    lob_elements=driver.find_elements(By.XPATH,lob_elements_xpaths)
    list={}
    for lob_element in lob_elements:
        action_click(driver,lob_element)
        count=driver.find_element(By.XPATH,count_xpaths).get_attribute('innerHTML')
        lob=lob_element.get_attribute('innerHTML')
        # if (lob == "Medicaid"):
        #     lob = "Medi-Cal"
        print(context_name+" "+year+" "+lob+" "+ count)
        list[lob]=int(count.replace("(", "").replace(")", "").replace(",",""))

    return list

#Call this to execute the Rollup Validation
def extract_context_data(driver, user):
    wait = WebDriverWait(driver, 50)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//span[text()="arrow_drop_down"]')))

    drop_down = driver.find_element(By.XPATH, '//span[text()="arrow_drop_down"]')
    action_click(driver, drop_down)
    time.sleep(5)

    group1_xpath = "//ul[@id='list-1']//li"
    group1_elements = driver.find_elements(By.XPATH, group1_xpath)
    print("Number of contexts present", len(group1_elements))

    for group1_s in group1_elements:
        print(group1_s.text)

    hierarchy_dropdown_open = True
    all_dictionary = {}
    other_dictionary = {}
    all_metric_dictionary = {}
    other_metric_dictionary = {}

    for i in range(1, len(group1_elements) + 1):
        if not hierarchy_dropdown_open:
            drop_down = driver.find_element(By.XPATH, '//span[text()="arrow_drop_down"]')
            action_click(driver, drop_down)
            time.sleep(3)
            hierarchy_dropdown_open = True

        group1_element_xpath = f"//ul[@id='list-1']//li[{i}]"
        action_click(driver, driver.find_element(By.XPATH, group1_element_xpath))

        apply_xpath = "//span[@class='context-menu-modal-apply']"
        action_click(driver, driver.find_element(By.XPATH, apply_xpath))
        hierarchy_dropdown_open = False
        ajax_preloader_wait(driver)

        context_name_xpath = "//span[@class='specific_most']"
        context_name = driver.find_element(By.XPATH, context_name_xpath).get_attribute("innerHTML")
        year = "2024"

        try:
            green_pill_xpath = "//a[@id='qt-filter-label']//span[2]"
            wait.until(EC.element_to_be_clickable((By.XPATH, green_pill_xpath)))
        except Exception as e:
            if ("access_denied" in driver.current_url):
                print(Fore.RED + f"Error : Access Denied observed for {context_name}"+ Style.RESET_ALL)
            continue

        temp_dict = extract_green_pill_info(driver, context_name, year)
        metric_temp_dict = extract_metric_lob_info(driver, context_name, year)

        if context_name == "All":
            all_dictionary = temp_dict
        else:
            for key, value in temp_dict.items():
                other_dictionary[key] = other_dictionary.get(key, 0) + value

        if context_name == "All":
            all_metric_dictionary = metric_temp_dict
        else:
            for key, value in metric_temp_dict.items():
                if key in other_metric_dictionary:
                    other_metric_dictionary[key] = [a + b for a, b in zip(other_metric_dictionary[key], value)]
                else:
                    other_metric_dictionary[key] = value

    print("ALL Metric Dictionary", all_metric_dictionary)
    print("Other Metric Dictionary", other_metric_dictionary)

    # Report directory that can
    link=create_html_report(user, all_dictionary, other_dictionary, "./reports")
    file_name=os.path.basename(link)
    report_path = f"./reports/{file_name}"
    append_comparison_report(all_metric_dictionary, other_metric_dictionary, report_path)
    return link

    print(all_dictionary)
    print(other_dictionary)