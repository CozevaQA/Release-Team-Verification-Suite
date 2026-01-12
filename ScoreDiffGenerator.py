import pickle
from datetime import datetime
from tkinter import *
from tkinter import ttk, filedialog

import os
import time
import traceback
import pandas as pd

from openpyxl.styles import Font, PatternFill
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
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


row_count = 0
chosen_client_list = []
chosen_date1_list = []
chosen_date2_list = []


def launchUI(customer_list):
    root = Tk()
    root.title("Scrollable Checklist")
    global date_list, row_count

    def add_row():
        global row_count
        row_count = row_count + 1

        date_list.append([])
        print("Row count:" + str(row_count))
        print("date_list_lenght: " + str(len(date_list)))
        selected_cust_var_list.append(StringVar())
        selected_cust_var_list[row_count].set("Select Customer")
        customer_drop_list.append(ttk.Combobox(root, textvariable=selected_cust_var_list[row_count], values=customer_list, state='readonly',
                                 style='TCombobox', width=35, height=35))


        date1_dropdown_var_list.append(StringVar())
        date1_dropdown_list.append(ttk.Combobox(root, textvariable=date1_dropdown_var_list[row_count], values=date_list[row_count], state='readonly',
                     style='TCombobox', width=35, height=35))


        date2_dropdown_var_list.append(StringVar())
        date2_dropdown_list.append(
            ttk.Combobox(root, textvariable=date2_dropdown_var_list[row_count], values=date_list[row_count], state='readonly',
                         style='TCombobox', width=35, height=35))
        customer_drop_list[row_count].grid(row=1+row_count, column=1)
        date1_dropdown_list[row_count].grid(row=1+row_count, column=2)
        date2_dropdown_list[row_count].grid(row=1+row_count, column=3)

    def confirm_button():
        global chosen_client_list, chosen_date1_list, chosen_date2_list
        chosen_client_list = []
        chosen_date1_list = []
        chosen_date2_list = []
        for client_var in selected_cust_var_list:
            chosen_client_list.append(client_var.get())
        for date_var in date1_dropdown_var_list:
            chosen_date1_list.append(date_var.get())
        for date_var in date2_dropdown_var_list:
            chosen_date2_list.append(date_var.get())

        print(chosen_client_list)
        print(chosen_date1_list)
        print(chosen_date2_list)

        root.destroy()


    def fetch_dates():
        altrowcounter = 0
        global date_list, date_list_primer
        chosen_client_list = []
        fetched_dates_list = []
        for client_var in selected_cust_var_list:
            chosen_client_list.append(client_var.get())

        print(chosen_client_list)
        for client in chosen_client_list:
            client_id = db.fetchCustomerID(client)
            path = os.path.join("assets\\ScoreFiles", str(client_id))
            files_in_path = [os.path.join(path, file) for file in os.listdir(path) if file.endswith('.pkl')]
            print(files_in_path)
            temp_date_list = []
            for file_path in files_in_path:
                file_path = file_path.replace("assets\\ScoreFiles\\"+str(client_id)+"\\", "")
                file_path = file_path.replace(".pkl", "")
                split_file_name = file_path.split("_")
                temp_date_list.append(split_file_name[1])
            temp_date_list = list(set(temp_date_list))
            date_list[altrowcounter] = date_list_primer + temp_date_list
            altrowcounter+=1

        for i, dropdown in enumerate(date1_dropdown_list):
            if i < len(date_list):
                dropdown['values'] = date_list[i]
            else:
                dropdown['values'] = []  # Clear the dropdown if no matching dates

        for i, dropdown in enumerate(date2_dropdown_list):
            if i < len(date_list):
                dropdown['values'] = date_list[i]
            else:
                dropdown['values'] = []  # Clear the dropdown if no matching dates






    def load_files():
        root_file_load = Tk()
        root_file_load.title("Load files")

        files_to_load = filedialog.askopenfilenames(
            title="Select Files",
            filetypes=(
                ("All Files", "*.*"),
                ("Text Files", "*.txt"),
                ("Python Files", "*.py")
            )
        )

        if files_to_load:
            print("Files selected:")
            for file in files_to_load:
                file_path = file
                file_name = file.split("/")[-1]  # Extract the file name from the path
                print(f"Full Path: {file_path}")
                print(f"File Name: {file_name}")

                #now, file processing and loading as pkl files

                current_dir = os.getcwd()

                path = os.path.join(current_dir+"/assets/", "ScoreFiles")
                isdir = os.path.isdir(path)
                if not isdir:
                    os.mkdir(path)

                file_params = file_name.split("_")
                for x in file_params:
                    print(x)

                path = os.path.join(path, str(file_params[0]))
                isdir = os.path.isdir(path)
                if not isdir:
                    os.mkdir(path)

                data = pd.read_csv(file_path, skiprows=2)
                print(data.head())
                print(f"Available columns in {file}: {list(data.columns)}")
                columns_to_keep = ["Measure Name", " Measure Abbreviation", " LOB Code", " Performance", " Numerator", " Denominator"]  # Replace with actual column names
                data = data[columns_to_keep]
                data_list = data.values.tolist()
                print(data_list)
                pklfilename = file_name.replace(".csv", ".pkl")
                with open(os.path.join(path,pklfilename), "wb") as cached_file:
                    pickle.dump(data_list, cached_file)

        else:
            print("No files selected.")

        root_file_load.destroy()

        #convert files to pkl here



        #fileprocessing here - pkl file will have 3d array.


        root_file_load.mainloop()


    row_count = 0
    add_row_button = ttk.Button(root, text="Add Row", command=add_row)
    add_row_button.grid(row=1, column=4)
    fetch_dates_button = ttk.Button(root, text="Fetch Dates", command=fetch_dates)
    fetch_dates_button.grid(row=1, column=5)
    load_files_button = ttk.Button(root, text="Load Files", command=load_files)
    load_files_button.grid(row=1, column=6)
    confirm_button = ttk.Button(root, text="Confirm", command=confirm_button)
    confirm_button.grid(row=1, column=7)
    customer_drop_list = []
    selected_cust_var_list = [StringVar()]
    selected_cust_var_list[row_count].set("Select Customer")

    customer_drop_list.append(ttk.Combobox(root, textvariable=selected_cust_var_list[row_count], values=customer_list, state='readonly',
                                 style='TCombobox', width=35, height=35))

    customer_drop_list[row_count].grid(row=1, column=1)
    global date_list
    print(date_list[0])
    date1_dropdown_var_list = [StringVar()]
    date1_dropdown_list = [
        ttk.Combobox(root, textvariable=date1_dropdown_var_list[row_count], values=date_list[row_count], state='readonly',
                     style='TCombobox', width=35, height=35)]

    date1_dropdown_list[row_count].grid(row=1, column=2)

    date2_dropdown_var_list = [StringVar()]
    date2_dropdown_list = [
        ttk.Combobox(root, textvariable=date2_dropdown_var_list[row_count], values=date_list[row_count], state='readonly',
                     style='TCombobox', width=35, height=35)]
    date2_dropdown_list[row_count].grid(row=1, column=3)

    root.mainloop()

def fetch_today_data(cust_id):
    driver = setups.driver_setup()
    setups.login_to_cozeva(cust_id)
    #time.sleep(40)
    fileset = []
    return fileset


def generate_diff():
    iterator = 0
    while iterator <= row_count:
        client = chosen_client_list[iterator]
        first_date = chosen_date1_list[iterator]
        first_date_fileset = []
        second_date = chosen_date2_list[iterator]
        second_date_fileset = []
        if first_date == "Today":
            first_date_fileset = fetch_today_data(db.fetchCustomerID(client))
        elif first_date == "Last Available Date":
            path = os.path.join("assets\\ScoreFiles", str(db.fetchCustomerID(client)))
            files_in_path = [os.path.join(path, file) for file in os.listdir(path) if file.endswith('.pkl')]
            temp_date_list = []
            for file_path in files_in_path:
                file_path = file_path.replace("assets\\ScoreFiles\\" + str(db.fetchCustomerID(client)) + "\\", "")
                file_path = file_path.replace(".pkl", "")
                split_file_name = file_path.split("_")
                temp_date_list.append(split_file_name[1])
            temp_date_list = list(set(temp_date_list))
            print(temp_date_list)
            temp_date_list_objects = [datetime.strptime(date, '%d%m%Y') for date in temp_date_list]
            today = datetime.today()
            valid_dates = [date for date in temp_date_list_objects if date <= today]
            latest_date = ""
            if valid_dates:
                latest_date = max(valid_dates)
                print("Latest date:", latest_date.strftime('%d%m%Y'))  # Format it back to string if needed
            else:
                print("No valid dates found.")

            #generate_fileset
            first_date_fileset.append(db.fetchCustomerID(client)+"_"+latest_date.strftime('%d%m%Y')+"_OFF_"+"2024"+".pkl")
            first_date_fileset.append(db.fetchCustomerID(client)+"_"+latest_date.strftime('%d%m%Y')+"_ON_"+"2024"+".pkl")
            first_date_fileset.append(db.fetchCustomerID(client)+"_"+latest_date.strftime('%d%m%Y')+"_OFF_"+"2025"+".pkl")
            first_date_fileset.append(db.fetchCustomerID(client)+"_"+latest_date.strftime('%d%m%Y')+"_ON_"+"2025"+".pkl")

            print(first_date_fileset)



        iterator+=1









    x=0


date_list_primer = ["Today", "Last Available Date", "Backup Only"]
date_list = [["Null"]]
export_download_path = os.path.join("c:\\VerificationReports", "DownloadDirectorydiff")
isdir = os.path.isdir(export_download_path)
if not isdir:
    os.mkdir(export_download_path)

export_download_path = os.path.join("assets", "cached_diffs")
isdir = os.path.isdir(export_download_path)
if not isdir:
    os.mkdir(export_download_path)

# File Naming convection = ClientID_date_MY_OFF/ON



# Sample data
# root = tk.Tk()
# root.title("Scrollable Checklist")
customer_list = db.getCustomerList()[1:len(db.getCustomerList())-1]
print(customer_list)
#date_list = []
launchUI(customer_list)
generate_diff()


# Example items; replace with db.getCustomerList() data
# items = [{'name': customer, 'last_backup_date': ""} for customer in customer_list]

# Create the scrollable checklist
# check_vars = create_scrollable_checklist(root, items)

# root.mainloop()
