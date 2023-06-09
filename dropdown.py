import os
import traceback
from tkinter import *

import openpyxl
from PIL import ImageTk, Image
from tkinter import ttk
import webbrowser
import ExcelProcessor as db

selected_cust = ""
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
    style.configure('TCombobox', fieldbackground=('readonly', 'green'), background=('readonly', 'white'), foreground=('readonly', 'black'))

    style.map('My.TButton', background=[('active', '#72B132')])

    def on_help():
        pdf_path = os.getcwd()
        pdf_file_path2 = "assets/RTVS Documentation.pdf"
        pdf_path = os.path.join(pdf_path, pdf_file_path2)
        webbrowser.open_new(pdf_path)

    def on_next():
        for index, element in enumerate(scriptselectorvar):
            scriptselectorvar[index] = scriptselectorvar[index].get()
        root.destroy()

        if scriptselectorvar[0] == 1:
            print("Running Measure Continuation Check")
        if scriptselectorvar[1] == 1:
            print("Running Provider Score Matching")
        if scriptselectorvar[2] == 1:
            print("Running Support level list Validation")

    cozeva_logo_image = ImageTk.PhotoImage(Image.open("assets/images/cozeva_logo.png").resize((280, 60)))
    help_icon_image = ImageTk.PhotoImage(Image.open("assets/images/help_icon.png").resize((15, 15)))

    logo_label = Label(root, image=cozeva_logo_image, background="white")
    logo_label.grid(row=1, column=0, padx=25, columnspan=4)


    please_select_label = Label(root, text="Post Computation Validation", background="white",
                                font=("Times New Roman", 15))
    please_select_label.grid(row=2, column=0, columnspan=4)

    customer_label = Label(root, text="Select customer", width="40", padx="40",background="white", font=("Times New Roman", 13))
    global selected_cust
    selected_cust = StringVar()
    selected_cust.set("Select Customer")
    customer_list = db.getCustomerList()  # vs.customer_list
    customer_drop = ttk.Combobox(root, textvariable=selected_cust, values=customer_list, state='readonly', style='TCombobox', width=35)
    customer_drop.set('Select Customer')

    #customer_label.grid(row=3, column=0, columnspan=4)
    customer_drop.grid(row=4, column=0, columnspan=4)


    help_button = ttk.Button(root, text="Help", command=on_help, image=help_icon_image,
                             compound="left", style='Configs.TButton')
    #help_button.grid(row=0, column=0, sticky='nw', padx=5, pady=5)

    scriptselectorvar = [IntVar(), IntVar(), IntVar()]

    # Add checkboxes
    checkbox1 = ttk.Checkbutton(root, text="Measure Continuity", variable=scriptselectorvar[0], style='CheckbuttonStyle.TCheckbutton')
    checkbox2 = ttk.Checkbutton(root, text="Provider Score Matching", variable=scriptselectorvar[1], style='CheckbuttonStyle.TCheckbutton')
    checkbox3 = ttk.Checkbutton(root, text="Support Level Lists", variable=scriptselectorvar[2], style='CheckbuttonStyle.TCheckbutton')

    checkbox1.grid(row=5, column=1, columnspan=4, padx=5, pady=5, sticky='w')
    checkbox2.grid(row=6, column=1, columnspan=4, padx=5, pady=5, sticky='w')
    checkbox3.grid(row=7, column=1, columnspan=4, padx=5, pady=5, sticky='w')

    # Add next button
    next_button = ttk.Button(root, text="Next", command=on_next, style='Next.TButton')
    next_button.grid(row=8, column=1, pady=20, columnspan=2)

    root.title("Post Computation Validation")
    root.iconbitmap("assets/icon.ico")
    root.mainloop()

try:
    master_gui()
except Exception as e:
    print(e)
    traceback.print_exc()
