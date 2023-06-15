import multiprocessing
import time
import os
import traceback
from tkinter import *
import ExcelProcessor as db

import openpyxl
from PIL import ImageTk, Image
from tkinter import ttk
import webbrowser
import os

# Yes, this is a convoluted way to maintain login info but im lazy and this is necessary
def check_and_create_file(file_path):
    if not os.path.exists(file_path):
        try:
            with open(file_path, 'w') as file:
                # Optionally, you can write some initial content to the file
                file.write(' ')
        except IOError:
            print("Login file exists")

# Usage
file_path = 'assets\loginInfo.txt'
check_and_create_file(file_path)

# This snippet will check if the logininfo file is empty. If it is, it will run first time setup
# then it will import the locator library
file = open(r"assets\loginInfo.txt", "r+")
file_content = str(file.read())
file.seek(0)
file.close()
if len(file_content) < 4:
    import FirstTimeSetup
    import variablestorage as locator
else:
    import variablestorage as locator

client_list = []


def rtvsmaster():
    # Store the directory the codebase is in for future calls
    code_directory = os.getcwd()
    try:
        root = Tk()
        root.configure(background='white')
        style = ttk.Style()
        style.theme_use('alt')
        style.configure('My.TButton', font=('Helvetica', 13, 'bold'), foreground='Black', background='#5a9c32',
                        padding=15, highlightthickness=0, height=1, width=25)
        style.configure('Configs.TButton', font=('Helvetica', 10, 'bold'), foreground='Black', background='#5a9c32',
                        highlightthickness=0)

        # style.configure('My.TButton', font=('American typewriter', 14), background='#232323', foreground='white')
        style.map('My.TButton', background=[('active', '#72B132')])

        def image_sizer(image_path):
            image_small = Image.open(image_path).resize((25, 25))

            return image_small

        # making image widgets
        first_time_setup_image = ImageTk.PhotoImage(image_sizer("assets/images/first_time_setup.png"))
        verification_suite_image = ImageTk.PhotoImage(image_sizer("assets/images/verification_suite.png"))
        hcc_validation_image = ImageTk.PhotoImage(image_sizer("assets/images/hcc_validation.png"))
        global_search_image = ImageTk.PhotoImage(image_sizer("assets/images/global_search.png"))
        task_ingestion_image = ImageTk.PhotoImage(image_sizer("assets/images/task_ingestion.png"))
        analytics_image = ImageTk.PhotoImage(image_sizer("assets/images/analytics.png"))
        slow_log_image = ImageTk.PhotoImage(image_sizer("assets/images/slow_log_trends.png"))
        pdf_printer_image = ImageTk.PhotoImage(image_sizer("assets/images/pdf_printer.png"))
        multi_role_image = ImageTk.PhotoImage(image_sizer("assets/images/Multi_role_access.png"))
        special_column_image = ImageTk.PhotoImage(image_sizer("assets/images/special_columns.png"))
        hospital_activity_image = ImageTk.PhotoImage(image_sizer("assets/images/hospital_activity.png"))
        supp_data_image = ImageTk.PhotoImage(image_sizer("assets/images/supp_data.png"))
        cozeva_logo_image = ImageTk.PhotoImage(Image.open("assets/images/cozeva_logo.png").resize((320, 71)))
        help_icon_image = ImageTk.PhotoImage(Image.open("assets/images/help_icon.png").resize((20, 20)))
        update_image = ImageTk.PhotoImage(Image.open("assets/images/update_image_2.png").resize((20, 20)))
        green_dot_image = ImageTk.PhotoImage(Image.open("assets/images/GreenDot.png").resize((10, 10)))
        red_dot_image = ImageTk.PhotoImage(Image.open("assets/images/RedDot.png").resize((10, 10)))
        orange_dot_image = ImageTk.PhotoImage(Image.open("assets/images/OrangeDot.png").resize((10, 10)))

        # Widgets+

        logo_label = Label(root, image=cozeva_logo_image, background="white")
        logo_label.grid(row=0, column=1)

        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)
        logo_label.grid(sticky="n")
        please_select_label = Label(root, text="Release Team Verification Suite", background="white",
                                    font=("Times New Roman", 15))
        please_select_label.grid(row=1, column=1)
        root.rowconfigure(1, weight=1)
        please_select_label.grid(sticky='n')

        # TRYING SOMETHING ELSE, HOPING THIS WORKS ITS 3 AM

        def on_first_time_setup():
            root.destroy()
            import FirstTimeSetup

        def on_verification_suite():
            root.destroy()
            import main

        def on_hcc_validation():
            root.destroy()
            import HCC_Validation_multi

        def on_global_search():
            root.destroy()
            import global_search

        def on_task_ingestion():
            root.destroy()
            import ProspectInjestHCC

        def on_analytics():
            root.destroy()
            import runner

        def on_slow_trends():
            root.destroy()
            import slowLogPlotter

        def on_role_access():
            root.destroy()

        def on_special_columns():
            root.destroy()
            import special_columns

        def on_hospital_activity():
            root.destroy()
            import Hospital_Activity

        def on_supp_data():
            root.destroy()
            # import Supplemental_data_alternate
            import secret_menu

        def on_submitbutton():
            global multi
            multi = 1
            for checkbox_index in range(0, len(customer_list)):
                if Checkbox_variables[checkbox_index].get() == 1:
                    client_list.append(db.fetchCustomerID(customer_list[checkbox_index]))

            print(client_list)

            root_overwatch.destroy()
            root.destroy()
            # import secret_menu


        def on_pdf_print():
            root.destroy()
            import printcareops

        def on_help():
            # root.destroy()
            pdf_path = os.getcwd()
            pdf_file_path2 = "assets/RTVS Documentation.pdf"
            pdf_path = os.path.join(pdf_path, pdf_file_path2)
            webbrowser.open_new(pdf_path)

        def on_update():
            import git
            import subprocess

            def has_new_commits():
                local_repo = os.getcwd()
                print(local_repo)
                subprocess.run(["git", "fetch"], check=True, cwd=local_repo, shell=True)
                result = subprocess.run(["git", "status", "-uno"], check=True, cwd=local_repo, shell=False,
                                        capture_output=True, text=True)
                output = result.stdout.strip()
                # print(status.stdout)
                if "Your branch is up to date" in output:
                    return False
                elif "Your branch is behind" in output:
                    return True
                return False

            if has_new_commits():
                update_button.configure(text="Updates Available")
            else:
                update_button.configure(text="No Updates")

        button_widgets = []
        button_widgets.append(
            ttk.Button(root, text="First time Setupwoooop", command=on_first_time_setup, image=first_time_setup_image,
                       compound="left", style='My.TButton'))
        button_widgets.append(
            ttk.Button(root, text="Verification Suite", command=on_verification_suite, image=verification_suite_image,
                       compound="left", style='My.TButton'))
        button_widgets.append(ttk.Button(root, text="HCC Validation Multi-Client", command=on_hcc_validation,
                                         image=hcc_validation_image, compound="left", style='My.TButton'))
        button_widgets.append(
            ttk.Button(root, text="Global Search", command=on_global_search, image=global_search_image, compound="left",
                       style='My.TButton'))
        button_widgets.append(ttk.Button(root, text="Task Ingestion(AWV)", command=on_task_ingestion,
                                         image=task_ingestion_image, compound="left", style='My.TButton'))
        button_widgets.append(
            ttk.Button(root, text="Analytics(Default)", command=on_analytics, image=analytics_image,
                       compound="left", style='My.TButton'))
        button_widgets.append(
            ttk.Button(root, text="Slow Log Trends", command=on_slow_trends, image=slow_log_image, compound="left",
                       style='My.TButton'))
        button_widgets.append(
            ttk.Button(root, text="PDF Print Validation", command=on_pdf_print, image=pdf_printer_image,
                       compound="left", style='My.TButton'))
        button_widgets.append(
            ttk.Button(root, text="Special Columns", command=on_special_columns, image=special_column_image,
                       compound="left", style='My.TButton'))
        button_widgets.append(ttk.Button(root, text="Hospital Activity (All Clients)", command=on_hospital_activity,
                                         image=hospital_activity_image, compound="left", style='My.TButton'))
        button_widgets.append(
            ttk.Button(root, text="Supplemental Data Addition", command=on_supp_data, image=supp_data_image,
                       compound="left", style='My.TButton'))
        button_widgets.append(
            ttk.Button(root, text="Overwatch", command=lambda: root_overwatch.deiconify(), image=hcc_validation_image,
                       compound="left", style='My.TButton'))

        help_button = ttk.Button(root, text="Help", command=on_help, image=help_icon_image,
                                 compound="left", style='Configs.TButton')
        update_button = ttk.Button(root, text="Check for Updates", command=on_update, image=update_image,
                                   compound="left", style='Configs.TButton')
        # widget counter to add the buttons in gridwise
        widget_counter = 0
        loopbreak = 0
        for i in range(2, 6):
            for j in range(3):
                try:
                    button_widgets[widget_counter].grid(row=i, column=j, padx=5, pady=5)
                except IndexError as e:
                    loopbreak = 1
                    break
                widget_counter += 1
            if loopbreak == 1:
                break

        help_button.grid(row=0, column=0, sticky='nw', padx=5, pady=5)
        update_button.grid(row=0, column=2, sticky='NE', padx=5, pady=5)

        # Chromeprofile multi threading

        chrome_profile_frame = Frame(root, background="white")
        Label(chrome_profile_frame, text="Chrome Profile Status", background="white",
              font=("Times New Roman", 15)).grid(row=0, column=0, columnspan=2)

        GUI_workbook = openpyxl.load_workbook('assets/profile_info.xlsx')
        GUI_sheet = GUI_workbook.active

        chrome_profile_info = []

        for row in GUI_sheet.iter_rows():
            row_data = []
            for cell in row:
                row_data.append(cell.value)
            chrome_profile_info.append(row_data)
        row_counter = 1
        for profile_row in chrome_profile_info:
            profile_name_label = Label(chrome_profile_frame, text=profile_row[1], background="white",
                                       font=("Times New Roman", 10))
            profile_name_label.grid(row=row_counter, column=0)
            profile_status_label = Label(chrome_profile_frame, text=profile_row[2], background="white",
                                         image=green_dot_image, compound="left", font=("Times New Roman", 10))
            profile_status_label.grid(row=row_counter, column=1, sticky="w", padx=10)
            row_counter += 1

            if profile_row[2] == "In Use":
                profile_status_label.configure(image=red_dot_image)
            if profile_row[1] == locator.free_chrome_profile:
                profile_status_label.configure(image=orange_dot_image, text=profile_row[2] + " (Current)")
                # test comment

        chrome_profile_frame.grid(row=1, rowspan=4, column=3, sticky="NE")

        root_overwatch = Toplevel(root)
        root_overwatch.title("Cozeva Overwatch")
        root_overwatch.iconbitmap("assets/icon.ico")
        root_overwatch.withdraw()
        customer_list = db.getCustomerList()
        # print(customer_list)
        Checkbox_variables = []
        for i in range(0, len(customer_list)):
            Checkbox_variables.append(IntVar())
        print(len(Checkbox_variables))
        Checkbox_widgets = []

        for i in range(0, len(customer_list)):
            Checkbox_widgets.append(Checkbutton(root_overwatch, text=customer_list[i], variable=Checkbox_variables[i],
                                                font=("Nunito Sans", 10)))
        submit_button = Button(root_overwatch, text="Submit", command=on_submitbutton, font=("Nunito Sans", 10))

        # add all checkboxes to a grid
        # practice_sidemenu_checkbox.grid(row=3, column=0, columnspan=5, sticky="w")
        for i in range(0, len(Checkbox_widgets)):
            if i <= 15:
                Checkbox_widgets[i].grid(row=i, column=0, sticky="w")
            elif 15 < i <= 30:
                Checkbox_widgets[i].grid(row=i - 15, column=2, sticky="w")
            elif 30 < i <= 45:
                Checkbox_widgets[i].grid(row=i - 30, column=3, sticky="w")
            elif 45 < i <= 60:
                Checkbox_widgets[i].grid(row=i - 45, column=4, sticky="w")
            submit_button.grid(row=0, column=4, sticky="w")

        root.title("Release Team Master Suite")
        root.iconbitmap("assets/icon.ico")
        # root.geometry("400x400+300+100")
        root.mainloop()

    except Exception as e:
        print(e)
        traceback.print_exc()
    finally:
        os.chdir(code_directory)
        file_location = "assets/profile_info.xlsx"
        chrome_profiles = openpyxl.load_workbook(file_location)
        chrome_profiles_sheet = chrome_profiles.active
        chrome_profile_available = False
        # Look for a row with an Available Chromeprofile name, Change it to In use and return the name
        for profile_index in range(1, 11):
            if str(chrome_profiles_sheet.cell(row=profile_index,
                                              column=2).value).strip() == locator.free_chrome_profile:
                chrome_profiles_sheet.cell(row=profile_index, column=3).value = 'Available'
                break

        chrome_profiles.save("assets/profile_info.xlsx")

def run_script(argument):
    def add_number_to_file(file_path, number):
        try:
            with open(file_path, 'a') as file:
                file.write(str(number) + '\n')
        except IOError as e:
            print(f"Error: {e}")

    file_path = 'assets\\overwatch_cache.txt'
    number = argument
    add_number_to_file(file_path, number)

    import multimain


if __name__ == '__main__':
    multiprocessing.freeze_support()
    multi = 0
    rtvsmaster()
    num_processes = len(client_list)

    if multi == 1:

        processes = []
        for arg in client_list:
            process = multiprocessing.Process(target=run_script, args=(arg,))
            process.start()
            processes.append(process)
            time.sleep(2)

        for process in processes:
            process.join()

        import clean_chrome_profiles
    else:
        print(multi)






