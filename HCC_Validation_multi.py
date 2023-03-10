import os
import time
import traceback

from openpyxl.styles import Font, PatternFill
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from tkinter import *
import setups
import logging
import ExcelProcessor as db
import context_functions as cf
import support_functions as sf
import setups as st
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import variablestorage as locator
from openpyxl import Workbook, load_workbook
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, \
    ElementClickInterceptedException, TimeoutException

ENV = 'PROD'
URL = ""
client_list = []
provider_count = 1
measurement_year = "2022"


def get_ID_list():
    root = Tk()
    customer_list = db.getCustomerList()
    #print(customer_list)
    Checkbox_variables = [None]*len(customer_list)
    for i in range(0, len(Checkbox_variables)):
        Checkbox_variables[i] = IntVar()
    print(len(Checkbox_variables))
    Checkbox_widgets = []



    def on_submitbutton():
        global provider_count, measurement_year
        provider_count = int(provider_count_entrybox.get().strip())
        measurement_year = str(measurement_year_entrybox.get().strip())

        for i in range(0,len(Checkbox_variables)):
            if Checkbox_variables[i].get() == 1:
                client_list.append(db.fetchCustomerID(customer_list[i]))
        print(client_list)
        root.destroy()


    #GenerateCheckboxesForall
    for i in range(0,len(customer_list)):
        Checkbox_widgets.append(Checkbutton(root, text=customer_list[i], variable=Checkbox_variables[i], font=("Nunito Sans", 10)))
    submit_button = Button(root, text="Submit", command=on_submitbutton, font=("Nunito Sans", 10))
    provider_count_entrybox = Entry(root)
    provider_measure_label = Label(root, text="Enter number of providers and MY", font=("Nunito Sans", 10))
    measurement_year_entrybox = Entry(root)
    #add all checkboxes to a grid
    #practice_sidemenu_checkbox.grid(row=3, column=0, columnspan=5, sticky="w")
    for i in range(0, len(Checkbox_widgets)):
        if i <= 15:
            Checkbox_widgets[i].grid(row=i, column=0, sticky="w")
        elif 15 < i <= 30:
            Checkbox_widgets[i].grid(row=i-15, column=2, sticky="w")
        elif 30 < i <= 45:
            Checkbox_widgets[i].grid(row=i-30, column=3, sticky="w")
        elif 45 < i <= 60:
            Checkbox_widgets[i].grid(row=i-45, column=4, sticky="w")
        submit_button.grid(row=0, column=4, sticky="w")
        provider_measure_label.grid(row=0, column=1, sticky="w")
        provider_count_entrybox.grid(row=0, column=2, sticky="w")
        measurement_year_entrybox.grid(row=0, column=3, sticky="w")
    root.title("Multi HCC Validation")
    root.iconbitmap("assets/icon.ico")
    root.mainloop()

if __name__ == '__main__':
    x=0

print("Hello World")
get_ID_list()

report_folder = os.path.join(locator.parent_dir,"HCC Multi Validation Reports")
isdir = os.path.isdir(report_folder)
if not isdir:
    os.mkdir(report_folder)
workbook_title = "HCC Multi Validation_"+sf.date_time()+".xlsx"

wb = Workbook()
ws = wb.active
ws.title = 'HCC_Validation ' + ENV
for ID in client_list:
    driver = setups.driver_setup()
    if ENV == 'CERT':
        setups.login_to_cozeva_cert(ID)
    elif ENV == 'STAGE':
        setups.login_to_cozeva_stage()
    elif ENV == "PROD":
        setups.login_to_cozeva(ID)
    else:
        print("ENV INVALID")
        exit(3)
    print("Run HCC Validation for "+str(ID)+ ",For providers: "+str(provider_count))

    cf.hccvalidation_multi(driver, ID, measurement_year, wb, provider_count, locator.parent_dir, "Cozeva Support", workbook_title)
    wb.save(report_folder + "\\" + workbook_title)
    driver.quit()




