import os
import time
import traceback
import webbrowser

import pygetwindow as gw
import pyautogui
from pywinauto import Desktop
from tkinter import *

import openpyxl
from PIL import ImageTk, Image
from tkinter import ttk

from openpyxl.styles import Font, PatternFill
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait

import setups
import logging
import ExcelProcessor as db
import context_functions as cf
import support_functions as sf
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import variablestorage as locator
from openpyxl import Workbook, load_workbook
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, \
    ElementClickInterceptedException, TimeoutException

ENV = 'CERT'
ID = "1000"
selected_cust = ""
sectionselectorvar = []
'''
    Section selector var section list
    0 = Patient Dashboard
    1 = Metric Specific patient list
    2 = Appointments Tab
    3 = Batches
    4 = Registries
    
'''



def master_gui():
    root = Tk()
    root.configure(background='white')
    style = ttk.Style()
    style.theme_use('alt')
    style.configure('My.TButton', font=('Helvetica', 13, 'bold'), foreground='Black', background='#5a9c32',
                    padding=15, highlightthickness=0, height=1, width=25)
    style.configure('Configs.TButton', font=('Helvetica', 8, 'bold'), foreground='Black', background='#5a9c32',
                    highlightthickness=0)
    style.configure('Next.TButton', font=('Helvetica', 10, 'bold'), foreground='Black', background='#5a9c32',
                    highlightthickness=0)
    style.configure('CheckbuttonStyle.TCheckbutton', font=('Helvetica', 13, 'bold'), foreground='Black', background='white')

    style.map('My.TButton', background=[('active', '#72B132')])

    def on_help():
        pdf_path = os.getcwd()
        pdf_file_path2 = "assets/RTVS Documentation.pdf"
        pdf_path = os.path.join(pdf_path, pdf_file_path2)
        webbrowser.open_new(pdf_path)

    def on_next():
        for index, section_element in enumerate(sectionselectorvar):
            sectionselectorvar[index] = sectionselectorvar[index].get()
        global ID
        ID = db.fetchCustomerID(selected_cust.get())
        print(selected_cust.get())
        print(ID)
        global Window_location
        Window_location = window_location_var.get()
        root.destroy()



    cozeva_logo_image = ImageTk.PhotoImage(Image.open("assets/images/cozeva_logo.png").resize((280, 60)))
    help_icon_image = ImageTk.PhotoImage(Image.open("assets/images/help_icon.png").resize((15, 15)))
    logo_label = Label(root, image=cozeva_logo_image, background="white")
    logo_label.grid(row=1, column=0, padx=50, columnspan=4)
    please_select_label = Label(root, text="PDF print validation", background="white", font=("Times New Roman", 15))
    please_select_label.grid(row=2, column=0, columnspan=4)

    global selected_cust
    selected_cust = StringVar()
    selected_cust.set("Select Customer")
    customer_list = db.getCustomerList()  # vs.customer_list
    customer_drop = OptionMenu(root, selected_cust, *customer_list)
    customer_drop.config(bg="#5a9c32", fg="black")

    #customer_label.grid(row=3, column=0, columnspan=4)
    customer_drop.grid(row=4, column=0, columnspan=4)


    help_button = ttk.Button(root, text="Help", command=on_help, image=help_icon_image,
                             compound="left", style='Configs.TButton')
    #help_button.grid(row=0, column=0, sticky='nw', padx=5, pady=5)

    global sectionselectorvar
    sectionselectorvar = [IntVar(), IntVar(), IntVar(), IntVar(), IntVar()]

    # Add checkboxes
    checkbox0 = ttk.Checkbutton(root, text="Patient Dashboard", variable=sectionselectorvar[0], style='CheckbuttonStyle.TCheckbutton')
    checkbox1 = ttk.Checkbutton(root, text="Metric Specific Patient List (WIP)", variable=sectionselectorvar[1], style='CheckbuttonStyle.TCheckbutton')
    checkbox2 = ttk.Checkbutton(root, text="Appointments Tab (WIP)", variable=sectionselectorvar[2], style='CheckbuttonStyle.TCheckbutton')
    checkbox3 = ttk.Checkbutton(root, text="Batches (WIP)", variable=sectionselectorvar[3], style='CheckbuttonStyle.TCheckbutton')
    checkbox4 = ttk.Checkbutton(root, text="Registries (WIP)", variable=sectionselectorvar[4], style='CheckbuttonStyle.TCheckbutton')


    checkbox0.grid(row=5, column=1, columnspan=4, padx=5, pady=5, sticky='w')
    checkbox1.grid(row=6, column=1, columnspan=4, padx=5, pady=5, sticky='w')
    checkbox2.grid(row=7, column=1, columnspan=4, padx=5, pady=5, sticky='w')
    checkbox3.grid(row=8, column=1, columnspan=4, padx=5, pady=5, sticky='w')
    checkbox4.grid(row=9, column=1, columnspan=4, padx=5, pady=5, sticky='w')

    Label(root, text="Which one is your laptop screen?", fg='red', background="white",
          font=("Times New Roman", 15)).grid(row=10, column=1, columnspan=4, padx=5, pady=5, sticky='w')

    window_location_label = Label(root, text="Select the screen for the testing window",
                                  font=("Nunito Sans", 10))
    window_location_var = IntVar()
    radiobutton_window_left = Radiobutton(root, text="Left", variable=window_location_var, value=1, background='white',
                                          font=("Nunito Sans", 10))
    radiobutton_window_right = Radiobutton(root, text="Right", variable=window_location_var, value=0, background='white',
                                           font=("Nunito Sans", 10))
    radiobutton_window_left.grid(row=11, column=1, padx=5, pady=5, sticky='w')
    radiobutton_window_right.grid(row=11, column=2, padx=5, pady=5, sticky='w')


    # Add next button
    next_button = ttk.Button(root, text="Next", command=on_next, style='Next.TButton')
    next_button.grid(row=12, column=1, pady=20, columnspan=2)

    root.title("PDF Print validation")
    root.iconbitmap("assets/icon.ico")
    root.mainloop()


def take_screenshot(window_title, output_path): #output_path needs to have an absolute path
    # Wait for the window to be available
    print("Screnshot function")
    desktop = Desktop(backend="uia")
    timeout = 10  # seconds
    end_time = time.time() + timeout
    while time.time() < end_time:
        try:
            window = desktop.window(title_re=window_title)
            if window.exists():
                print("Window Found")
                break
        except Exception:
            pass
        time.sleep(0.5)
    else:
        print(f"Could not find the window with title matching '{window_title}'")

    # Bring the window to the front
    window.set_focus()

    # Take the screenshot
    screenshot = pyautogui.screenshot(region=(
        window.rectangle().left, window.rectangle().top, window.rectangle().width(), window.rectangle().height()))
    screenshot.save(output_path)

    time.sleep(2)
    window.set_focus()
    pyautogui.press('esc')
    time.sleep(2)


def patient_dashboard_print(driver):
    #Add the logic to navigate to a patient dashboard from CS account registry
    driver.get("https://cert.cozeva.com/patient_detail/12STMW0?tab_type=CareOps&cozeva_id=12STMW0&patient_id=8717219&cozeva_id=12STMW0&session=YXBwX2lkPXJlZ2lzdHJpZXMmY3VzdElkPTEwMDAmZG9jdG9yc1BlcnNvbklkPTYxOTcxNjgmZG9jdG9yX3VpZD02MTc2OTUyJnBheWVySWQ9MTAwMCZxdWFydGVyPTIwMjMtMTItMzEmaG9tZT1ZWEJ3WDJsa1BYSmxaMmx6ZEhKcFpYTW1ZM1Z6ZEVsa1BURXdNREFtY0dGNVpYSkpaRDB4TURBd0ptOXlaMGxrUFRFd01EQSZmaWx0ZXJfb3JnX2lkPQ%3D%3D&first_load=1")
    sf.ajax_preloader_wait(driver)

    driver.find_element(By.CLASS_NAME, "patient_print_options").click()
    time.sleep(2)

    dropdown_elements = driver.find_element(By.ID, "patient_print_dropdown").find_elements(By.TAG_NAME, "li")
    print_options = ["Print Careops Summary", "Print HCC Confirm/Disconfirm", "Print Careops", "Print Med Adherence", "Print HCC"]
    for element in dropdown_elements:
        print(element.text)
    file_counter = 1
    for element in dropdown_elements:
        if element.text in print_options:
            try:

                print("Clicking on "+element.text)
                printss_filename = str(element.text).replace(" ", "_") + ".png"
                bad_chars = [';', ':', '|', ' ', '/', '\\']
                for i in bad_chars:
                    printss_filename = printss_filename.replace(i, '_')
                print(printss_filename)
                #element.click()
                #driver.execute_script("arguments[0].click();", element)
                webdriver.ActionChains(driver).move_to_element(element).click(element).perform()
                #print(driver.window_handles)
                print("Clicked on "+element.text)
                time.sleep(4)
                print("Waiting done")
                #driver.execute_script("window.focus();")

                take_screenshot('.*Print', report_folder + "\\" + printss_filename)
                file_counter += 1


                driver.find_element(By.CLASS_NAME, "patient_print_options").click()
                time.sleep(2)


            except Exception as e:
                traceback.print_exc()


        time.sleep(2)


report_folder = os.path.join(locator.parent_dir, "PDF Print Reports")
isdir = os.path.isdir(report_folder)
if not isdir:
    os.mkdir(report_folder)
report_folder = os.path.join(report_folder, str(sf.date_time()))
isdir = os.path.isdir(report_folder)
if not isdir:
    os.mkdir(report_folder)
Window_location = 1
try:
    master_gui()
except Exception as e:
    print(e)
    traceback.print_exc()

driver = setups.driver_setup()
print(Window_location)
if Window_location == 1:
    driver.set_window_position(-1000, 0)
elif Window_location == 0:
    driver.set_window_position(1000, 0)
driver.maximize_window()


if ENV == 'CERT':
    setups.login_to_cozeva_cert(ID)
elif ENV == 'STAGE':
    setups.login_to_cozeva_stage()
elif ENV == "PROD":
    setups.login_to_cozeva(ID)
else:
    print("ENV INVALID")
    exit(3)

sf.ajax_preloader_wait(driver)

if sectionselectorvar[0] == 1:
    print("Running patient dashboard print validation")
    patient_dashboard_print(driver)
if sectionselectorvar[1] == 1:
    print("Running MSPL prints validation")
if sectionselectorvar[2] == 1:
    print("Running appointments tab print validation")
if sectionselectorvar[3] == 1:
    print("Running batches print validation")
if sectionselectorvar[4] == 1:
    print("Running registries print validation")




