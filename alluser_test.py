# All view scripting NOT all user
#
# 1. User will be predefined
# 2. Script will masquerade to the user -
# 3. Checks to apply (Check time taken for all click actions)
# 	-Check if account is All user (either through context name, maybe some other indicator) - When is an account considered all user? is it ONLY when the account contains other linked accounts?
# 	-Predefined list of sidebar options, click on each option on the sidebar, report blank lists, report missing options/new unexpected options (will this depend on the main user's main role)
# 	-Add up patient counts on each context registry - Match vs all view patient count
# 	-Export options in All view - Export current Registry, All registries, filtered Registries
# 	-Verify resources dropdown on each link
# 	-Check each apptray option, list ones greyed out, and ones that loaded

# revised scope - Predefined user of each role - Sidebar options, Support level metric specific lists, apptray options access, global search, All user session  -
# Review columns, document preview, screenshot if fail(maybe no) - NPI search - hardcode filter search on lists - GSD@
# Keep separate - NPI - Pickup from provider list.
#future scope - Include as part of daily validSummation of numerator/denominator roll up
import pytest
import allure
import os
import random
from statistics import mean
import sys
import math
import time
import traceback
import csv
from os import listdir
from os.path import isfile, join
import csv
import time
from datetime import date
from datetime import datetime
import base64
from tkinter import messagebox, ttk

import pandas as pd
from PIL import Image as img
from PIL import ImageTk
import PIL
from pytest_assume.plugin import assume
from selenium import *
from selenium import webdriver
from selenium.webdriver.common import keys
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import *
from selenium.webdriver.support.wait import WebDriverWait
from openpyxl import Workbook
from openpyxl.styles import PatternFill, NamedStyle, Border
from urllib.parse import urlparse, parse_qs
from openpyxl.styles import *
import support_functions as sf
import setups
import variablestorage as locator
import ExcelProcessor as db
from tkinter import *

ENV = "CERT"
predefined_user = 'CozevaQA_Prov_Alt'
client_id = '1500'
driver = setups.driver_setup()

if ENV == 'CERT':
    setups.login_to_cozeva_cert(client_id)
elif ENV == 'STAGE':
    setups.login_to_cozeva_stage()
elif ENV == "PROD":
    setups.login_to_cozeva(client_id)
else:
    print("ENV INVALID")
    exit(3)

setups.login_to_user(predefined_user)
print("Landing page= " + driver.title)
sf.ajax_preloader_wait(driver)
sf.skip_intro(driver)
print("Masq to user successful")

@allure.title('''Check if account is All user (either through context name, maybe some other indicator) - 
              When is an account considered all user? is it ONLY when the account contains other linked accounts?''')
def test_check_login_is_all_user():
    current_context_name = driver.find_element(By.XPATH, locator.xpath_context_Name).text
    assume(current_context_name == "All User","Check if account is All user")

@allure.title('''# 	-Predefined list of sidebar options, click on each option on the sidebar, report blank lists, 
                report missing options/new unexpected options (will this depend on the main user's main role)''')
def test_sidebar_options():
    with allure.step("Clicking on Sidebar Option: Supplemental Data"):
        registry_url = driver.current_url
        driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()
        links = driver.find_elements_by_xpath(locator.xpath_menubar_Item_Link)
        names = driver.find_elements_by_xpath(locator.xpath_menubar_Item_Link_Name)
        print("entered supplimental data")

        for link in range(len(links)) and range(len(names)):
            if names[link].text == "Supplemental Data":
                print("entered supplimental data if block")
                links[link].click()
                start_time = time.perf_counter()
                sf.ajax_preloader_wait(driver)
                total_time = time.perf_counter() - start_time
                current_url = driver.current_url
                # Need to check that the page has opened properly
                table_name = driver.find_element(By.CLASS_NAME, "table_header").text
                print("CURRENT TABLE NAME"+table_name)

                assume(table_name == "Supplemental Datao", "Step 1: Supplemental Data")

                driver.get(registry_url)
                sf.ajax_preloader_wait(driver)
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
                links, names = [], []

    with allure.step("Clicking on Sidebar Option: HCC Chart List"):
        registry_url = driver.current_url
        driver.find_element_by_xpath(locator.xpath_side_nav_SlideOut).click()
        links = driver.find_elements_by_xpath(locator.xpath_menubar_Item_Link)
        names = driver.find_elements_by_xpath(locator.xpath_menubar_Item_Link_Name)

        for link in range(len(links)) and range(len(names)):
            if names[link].text == "HCC Chart List":
                print("entered supplimental data if block")
                links[link].click()
                start_time = time.perf_counter()
                sf.ajax_preloader_wait(driver)
                total_time = time.perf_counter() - start_time
                current_url = driver.current_url
                # Need to check that the page has opened properly
                table_name = driver.find_element(By.CLASS_NAME, "table_header").text

                assume(table_name == "HCC Chart List", "Step 2: HCC Chart List")

                driver.get(registry_url)
                sf.ajax_preloader_wait(driver)
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH, locator.xpath_filter_measure_list)))
                links, names = [], []









