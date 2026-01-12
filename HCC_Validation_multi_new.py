import pickle
import tkinter as tk

import openpyxl
from PIL import ImageTk, Image
from tkinter import ttk
import setups
import logging
import ExcelProcessor as db
import context_functions as cf
import support_functions as sf
import variablestorage as locator

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook, load_workbook
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, \
    ElementClickInterceptedException, TimeoutException

selected_cust = "1500"
provider_count = 1
selected_my = "2025"
def master_gui():
    root = tk.Tk()
    root.configure(background='white')
    style = ttk.Style()
    style.theme_use('alt')
    style.configure('My.TButton', font=('Helvetica', 13, 'bold'), foreground='Black', background='#5a9c32',
                    padding=15, highlightthickness=0, height=1, width=25)
    style.configure('Configs.TButton', font=('Helvetica', 8, 'bold'), foreground='Black', background='#5a9c32',
                    highlightthickness=0)
    style.configure('Next.TButton', font=('Helvetica', 10, 'bold'), foreground='Black', background='#5a9c32',
                    highlightthickness=0)
    style.configure('CheckbuttonStyle.TCheckbutton', font=('Helvetica', 15, 'bold'), foreground='Black', background='white')
    style.configure('RadiobuttonStyle.TRadiobutton', font=('Helvetica', 10, 'bold'), foreground='black', background='white')
    style.configure('FrameStyle.Tframe', foreground='black', background='white')

    style.map('My.TButton', background=[('active', '#72B132')])
    def on_next():
        #what should happen here is based on the year selected, make a pkl of MY, Prov count, and client ID.
        global selected_cust, selected_my, provider_count
        selected_cust = selected_cust_var.get()
        selected_my = selected_year_var.get()
        provider_count = provider_count_var.get()

        print(selected_cust +" "+str(selected_my)+" "+str(provider_count))
        root.destroy()

    selected_year_var = tk.IntVar(value=2023)
    MY_frame = tk.Frame(root, background='white')
    # Generate radio buttons horizontally for years 2023–2028
    for year in range(2023, 2029):
        rb = ttk.Radiobutton(MY_frame, text=str(year), style='RadiobuttonStyle.TRadiobutton', variable=selected_year_var, value=year)
        rb.pack(side="left", padx=10)


    cozeva_logo_image = ImageTk.PhotoImage(Image.open("assets/images/cozeva_logo.png").resize((320, 74)))
    help_icon_image = ImageTk.PhotoImage(Image.open("assets/images/help_icon.png").resize((15, 15)))
    logo_label = tk.Label(root, image=cozeva_logo_image, background="white")
    logo_label.grid(row=1, column=0, padx=50, columnspan=4)
    please_select_label = tk.Label(root, text="Please Select Client, Provider count and Measurement Year", background="white", font=("Times New Roman", 13))
    please_select_label.grid(row=2, column=0, columnspan=4)

    global selected_cust
    selected_cust_var = tk.StringVar()
    selected_cust_var.set("Select Customer")
    customer_list = db.getCustomerList()  # vs.customer_list
    customer_drop = ttk.Combobox(root, textvariable=selected_cust_var, values=customer_list, state='readonly', style='TCombobox', width=35, height=35)
    #customer_drop.config(bg="#5a9c32", fg="black")

    global provider_count
    provider_count_var = tk.IntVar()
    provider_count_var.set(1)
    provider_drop = ttk.Combobox(root, textvariable=provider_count_var, values=['1','2','3','4','5'], state='readonly', style='TCombobox', width=20)






    #customer_label.grid(row=3, column=0, columnspan=4)
    customer_drop.grid(row=4, column=0, columnspan=2)
    provider_drop.grid(row=4, column=2, columnspan=2)

    MY_frame.grid(row=5, column=0, columnspan=4, padx=30)

    # Add next button
    next_button = ttk.Button(root, text="Next", command=on_next, style='Next.TButton')
    next_button.grid(row=12, column=1, pady=20, columnspan=2)

    root.title("PDF Print validation")
    root.iconbitmap("assets/icon.ico")
    root.mainloop()

master_gui()

with open("assets\\hcc_data.pkl", 'wb') as hcc_file:
    pickle.dump([db.fetchCustomerID(selected_cust),selected_my, provider_count], hcc_file)

# blended = 23,24
# rest = 25+
if int(selected_my) <=2024:
    with open("HCC_Blended_RTVS.py") as hcc_code_blended:
        exec(hcc_code_blended.read())
elif int(selected_my) >=2024:
    with open("HCC_v28_only_RTVS.py") as hcc_code_v28:
        exec(hcc_code_v28.read())


