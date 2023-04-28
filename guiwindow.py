import time
import traceback
from os import getcwd
from tkinter import *

import variablestorage as locator
from selenium import webdriver

import ExcelProcessor as db
import Schema_processor as sp

checklist = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
roleset = {"Cozeva Support": "99999"}
verification_specs = ["Name", 9999, "Onshore", roleset, checklist]
env = "STAGE"
headlessmode = 0
Window_location = 1 #1 = left, 0 = Right
grid_row = 2


'''This module is used to store scripts that flash and record GUI input'''

""" 1. Cache on/off check (analytics)
    2. first page - choice between cozeva/analytics
    3. Check analytics through cs2
    4. analytics based on cust,region,panel(cohort analyser/Quality)
    5. Full cust level checkbox (Redshift transfer checkbox. Do both)"""


def launchgui():
    """ README

    This function will return a list containing the specifications of the test to be run
    roleset is a dict that will return roles and usernames as a key value pair
    The verification_specs variable is a list that will be returned
    index tracking:
    0 = Customer name (Default value = "Name")
    1 = Customer ID (Default value = "9999")
    2 = offshore/Onshore (Default value = "Onshore")
    3 = Roleset with username(Default value = ["cs2:default"])
    4 = boolean list of things to be tested (Ex, Support registry sidemenu, Practice context navigations)

    index tracking of boolean list
    0 = Sidemenu (Support level)
    1 = Sidemenu (Practice level)
    2 = Sidemenu (Provider level)
    3 = Patient Dashboard
    4 = Provider Registry
    5 = Practice Registry
    6 = Support Registry
    7 = Global Search
    8 = Provider's MSPL
    9 = Time Capsule
    10 = Secure Messaging
    11 = Analytics
    12 = Accordion Verification
    13 = Analytics Full
    14 = Measure Validation LOBwise
    15 = Metric Specific Provider List
    16 = Cozeva Market sheet( Check default LoB and default CE from mastersheet, Customer name too and CE on/off score change)
    17 = Check patient Medications
    18 = Mark as pending
    19 = Apptray access check
    20 = Training/Resources tab
    21 = Userlist and User creation dropdown
    22 = Banner/announcement
    23 = CE Toggle States
    24 = Sticket and contact Log
    25 = Patient Dashboard and Timeline Redirection
    26 = Coding tool
    27 = Denominator Eligblity for CBP/HbA1c
    28 = Group1/2 sidemenus
    29 = HCC Validation


    """


    #creating the root canvas
    root = Tk()

    #Creating Principle frames
    input_frame_1 = Frame(root)
    input_frame_2 = Frame(root)
    input_frame_3 = Frame(root)
    Analytics_frame = Frame(root)
    new_launch_frame = Frame(root)


    #local function to delete frame content
    def remove_frames():
        input_frame_1.pack_forget()
        input_frame_2.pack_forget()
        input_frame_3.pack_forget()
        Analytics_frame.pack_forget()
        new_launch_frame.pack_forget()


    def frame1next():
        global env, headlessmode, Window_location
        if env_selector_var.get() == 0:
            env = "STAGE"
        headlessmode = headless_selector_var.get()
        verification_specs[0] = selected_cust.get()
        verification_specs[1] = db.fetchCustomerID(verification_specs[0])
        print(window_location_var.get())
        Window_location = window_location_var.get()
        print(Window_location)
        if Checkbox_analytics_var.get() == 1:
            checklist[11] = 1
        if Checkbox_analytics_var.get() == 1 and Checkbox_cozeva_var.get() == 0 and config_var.get() == 0:
            root.destroy()
            return
        if config_var.get() == 0 and Checkbox_cozeva_var.get() == 1:
            roleset.update(db.getDefaultUserNames(selected_cust.get()))
            verification_specs[3] = roleset
            if len(roleset) > 1:
                verification_specs[2] = "Offshore"
        if config_var.get() == 0 and app_config_var.get() != 1:
            for i in range(11):
                checklist[i] = 1
            checklist[12]=1
            verification_specs[4] = checklist
            root.destroy()
            return

        if config_var.get() == 1 and Checkbox_cozeva_var.get() == 1:
            remove_frames()
            default_usernames = db.getDefaultUserNames(selected_cust.get())
            input_frame_2.pack()
            insert_into_entrybox(default_usernames)
        if config_var.get() == 2 and app_config_var.get() != 1:
            remove_frames()
            new_launch_frame.pack()
        if app_config_var.get() == 1 and config_var.get() == 1:
            remove_frames()
            checklist[13] = 1
            Analytics_frame.pack()
            buildExistingSchema()
        if app_config_var.get() == 1 and config_var.get() == 0:
            checklist[13] = 1
            root.destroy()
            return





    def frame2next():
        create_roleset()
        verification_specs[3] = roleset
        if len(roleset) > 1:
            verification_specs[2] = "Offshore"
        remove_frames()
        input_frame_3.pack()


    def frame2prev():
        delete_from_entrybox()
        remove_frames()
        input_frame_1.pack()

    def frame3next():
        create_checklist()
        root.destroy()
        verification_specs[4] = checklist
        #for i in checklist:
        #    print(i)
        #return verification_specs

    def frame3prev():
        remove_frames()
        input_frame_2.pack()

    def nlnext():
        create_NL_checklist()
        verification_specs[4] = checklist
        verification_specs[1] = str(NL_custID.get())
        verification_specs[0] = "NC_"+verification_specs[1]
        root.destroy()

    def delete_from_entrybox():
        limited_cozeva_support_textbox.delete(0, 'end')
        customer_support_textbox.delete(0, 'end')
        regional_support_textbox.delete(0, 'end')
        office_admin_practice_textbox.delete(0, 'end')
        office_admin_provider_textbox.delete(0, 'end')
        provider_textbox.delete(0, 'end')



    def insert_into_entrybox(usernames_for_insert):
        for i in usernames_for_insert:
            if i == "Cozeva Support":
                cozeva_support_checkbox.select()
            if i == "Limited Cozeva Support":
                limited_cozeva_support_textbox.insert(0, usernames_for_insert[i])
                limited_cozeva_support_checkbox.select()
            if i == "Customer Support":
                customer_support_textbox.insert(0, usernames_for_insert[i])
                customer_support_checkbox.select()
            if i == "Regional Support":
                regional_support_textbox.insert(0, usernames_for_insert[i])
                regional_support_checkbox.select()
            if i == "Office Admin Practice Delegate":
                office_admin_practice_textbox.insert(0, usernames_for_insert[i])
                office_admin_practice_checkbox.select()
            if i == "Office Admin Provider Delegate":
                office_admin_provider_textbox.insert(0, usernames_for_insert[i])
                office_admin_provider_checkbox.select()
            if i == "Provider":
                provider_textbox.insert(0,usernames_for_insert[i])
                provider_checkbox.select()

    def create_roleset():
        if limited_cozeva_support_var.get() == 1:
            roleset.update({"Limited Cozeva Support":limited_cozeva_support_textbox.get().strip()})
        if customer_support_var.get() == 1:
            roleset.update({"Customer Support": customer_support_textbox.get().strip()})
        if regional_support_var.get() == 1:
            roleset.update({"Regional Support": regional_support_textbox.get().strip()})
        if office_admin_practice_var.get() == 1:
            roleset.update({"Office Admin Practice Delegate": office_admin_practice_textbox.get().strip()})
        if office_admin_provider_var.get() == 1:
            roleset.update({"Office Admin Provider Delegate": office_admin_provider_textbox.get().strip()})
        if provider_var.get() == 1:
            roleset.update({"Provider": provider_textbox.get().strip()})

    def create_checklist():
        checklist[0] = support_sidemenu_var.get()
        checklist[1] = practice_sidemenu_var.get()
        checklist[2] = provider_sidemenu_var.get()
        checklist[3] = patient_dashboard_var.get()
        checklist[4] = provider_registry_var.get()
        checklist[5] = practice_registry_var.get()
        checklist[6] = support_registry_var.get()
        checklist[7] = global_search_var.get()
        checklist[8] = provider_mspl_var.get()
        checklist[9] = time_capsule_var.get()
        checklist[10] = secure_messaging_var.get()
        checklist[12] = accordion_verification_var.get()
        checklist[29] = hcc_validation_var.get()

    def create_NL_checklist():
        if all_navigation_var.get() == 1:
            checklist[0] = 1
            checklist[1] = 1
            checklist[2] = 1
            checklist[4] = 1
            checklist[5] = 1
            checklist[28] = 1
        checklist[14] = LoB_Measure_var.get()
        checklist[15] = provider_tab_var.get()
        checklist[16] = market_sheet_var.get()
        checklist[17] = patient_medication_var.get()
        checklist[18] = MaP_var.get()
        checklist[19] = apptray_var.get()
        checklist[20] = train_resource_var.get()
        checklist[21] = userlist_var.get()
        checklist[22] = banner_announce_var.get()
        checklist[23] = ce_toggle_var.get()
        checklist[12] = NL_accordion_verification_var.get()
        checklist[7] = NL_global_search_var.get()
        checklist[24] = sticket_var.get()
        checklist[25] = NL_patient_dashboard_var.get()
        checklist[26] = coding_tool_var.get()
        checklist[27] = denom_eligibility_var.get()
        checklist[29] = NL_hcc_validation_var.get()

    def select_all():
        if select_var.get() == 1:
            support_sidemenu_checkbox.select()
            practice_sidemenu_checkbox.select()
            provider_sidemenu_checkbox.select()
            patient_dashboard_checkbox.select()
            provider_registry_checkbox.select()
            practice_registry_checkbox.select()
            support_registry_checkbox.select()
            global_search_checkbox.select()
            provider_mspl_checkbox.select()
            time_capsule.select()
            secure_messaging_checkbox.select()
            accordion_verification_checkbox.select()
            hcc_validation_checkbox.select()
        elif select_var.get() == 0:
            support_sidemenu_checkbox.deselect()
            practice_sidemenu_checkbox.deselect()
            provider_sidemenu_checkbox.deselect()
            patient_dashboard_checkbox.deselect()
            provider_registry_checkbox.deselect()
            practice_registry_checkbox.deselect()
            support_registry_checkbox.deselect()
            global_search_checkbox.deselect()
            provider_mspl_checkbox.deselect()
            time_capsule.deselect()
            secure_messaging_checkbox.deselect()
            accordion_verification_checkbox.deselect()
            hcc_validation_checkbox.select()

    def NL_select_all():
        if NL_select_var.get() == 1:
            all_navigation_checkbox.select()
            #LoB_Measure_checkbox.select()
            provider_tab_checkbox.select()
            market_sheet_checkbox.select()
            patient_medication_checkbox.select()
            MaP_checkbox.select()
            apptray_checkbox.select()
            train_resource_checkbox.select()
            #userlist_checkbox.select()
            #banner_announce_checkbox.select()
            ce_toggle_checkbox.select()
            NL_accordion_verification_checkbox.select()
            NL_global_search_checkbox.select()
            sticket_checkbox.select()
            NL_patient_dashboard_checkbox.select()
            coding_tool_checkbox.select()
            #denom_eligibility_checkbox.select()
            NL_hcc_validation_checkbox.select()
        elif NL_select_var.get() == 0:
            all_navigation_checkbox.deselect()
            #LoB_Measure_checkbox.deselect()
            provider_tab_checkbox.deselect()
            market_sheet_checkbox.deselect()
            patient_medication_checkbox.deselect()
            MaP_checkbox.deselect()
            apptray_checkbox.deselect()
            train_resource_checkbox.deselect()
            #userlist_checkbox.deselect()
            #banner_announce_checkbox.deselect()
            ce_toggle_checkbox.deselect()
            NL_accordion_verification_checkbox.deselect()
            NL_global_search_checkbox.deselect()
            sticket_checkbox.deselect()
            NL_patient_dashboard_checkbox.deselect()
            coding_tool_checkbox.deselect()
            #denom_eligibility_checkbox.deselect()
            NL_hcc_validation_checkbox.deselect()

    def cozeva_radio():
        Checkbox_cozeva.config(state="active")
        Checkbox_analytics.config(state="active")
        Checkbox_cozeva.select()
        Checkbox_analytics.select()
        customer_drop.config(state='active')
        radiobutton_env_stage.config(state='active')
        #customer_drop.set('Customer')

    def analytics_radio():
        Checkbox_cozeva.deselect()
        Checkbox_analytics.deselect()
        Checkbox_cozeva.config(state="disabled")
        Checkbox_analytics.config(state="disabled")
        customer_drop.config(state='disabled')



    #widgets for frame 1
    warninglabel = Label(input_frame_1,
                         text="ERROR OCCURED. CLOSE YOUR TEST WINDOW AND/OR UPDATE YOUR CHROME BROWSER/CHROME DRIVER", fg="red", font=("Nunito Sans", 12))
    customer_label = Label(input_frame_1, text="Select customer", width="40", padx="40", font=("Nunito Sans", 10))
    global selected_cust
    selected_cust = StringVar()
    selected_cust.set("Customer")
    customer_list = db.getCustomerList() #vs.customer_list
    customer_drop = OptionMenu(input_frame_1, selected_cust, *customer_list)
    context_label = Label(input_frame_1, text="Context", font=("Nunito Sans", 10))
    config_label = Label(input_frame_1, text="Configuration", font=("Nunito Sans", 10))
    Checkbox_cozeva_var = IntVar()
    Checkbox_cozeva = Checkbutton(input_frame_1, text="Cozeva Verification", variable=Checkbox_cozeva_var, font=("Nunito Sans", 10))
    Checkbox_analytics_var = IntVar()
    Checkbox_analytics = Checkbutton(input_frame_1, text="Analytics Verification", variable=Checkbox_analytics_var, font=("Nunito Sans", 10))
    config_var = IntVar()
    radiobutton_default = Radiobutton(input_frame_1, text="Default Configuration", variable=config_var, value=0, font=("Nunito Sans", 10))
    radiobutton_illchoose = Radiobutton(input_frame_1, text="I'll choose", variable=config_var, value=1, font=("Nunito Sans", 10))
    radiobutton_newcustomer = Radiobutton(input_frame_1, text="Customer Launch config", variable=config_var, value=2, font=("Nunito Sans", 10))
    window_location_label = Label(input_frame_1, text="Select the screen for the testing window", font=("Nunito Sans", 10))
    window_location_var = IntVar()
    radiobutton_window_left = Radiobutton(input_frame_1, text="Left", variable=window_location_var, value=1, font=("Nunito Sans", 10))
    radiobutton_window_right = Radiobutton(input_frame_1, text="Right", variable=window_location_var, value=0, font=("Nunito Sans", 10))
    app_config_var = IntVar()
    config_cozeva = Radiobutton(input_frame_1, text="Cozeva Verification", variable=app_config_var, value=0,
                                command=cozeva_radio)
    config_analytics = Radiobutton(input_frame_1, text="Analytics Verification(Full)", variable=app_config_var, value=1,
                                   command=analytics_radio)
    nextbutton1 = Button(input_frame_1, text="Lock choices and Proceed", command=frame1next, font=("Nunito Sans", 10))
    env_selector_var = IntVar()
    radiobutton_env_label = Label(input_frame_1, text="Select Environment", font=("Nunito Sans", 10))
    radiobutton_env_stage = Radiobutton(input_frame_1, text="STAGE", variable=env_selector_var, value=0, font=("Nunito Sans", 10))

    headless_selector_var = IntVar()
    radiobutton_headless_label = Label(input_frame_1, text="Select Headless mode Yes/No", font=("Nunito Sans", 10))
    radiobutton_headless_yes = Radiobutton(input_frame_1, text="Yes", variable=headless_selector_var, value=1, font=("Nunito Sans", 10))
    radiobutton_headless_no = Radiobutton(input_frame_1, text="No", variable=headless_selector_var, value=0, font=("Nunito Sans", 10))
    #widgets for frame 2
    frame2_info_label = Label(input_frame_2, text="Select roles and enter username(Leave blank for default)", font=("Nunito Sans", 10))
    cozeva_support_var = IntVar()
    limited_cozeva_support_var = IntVar()
    customer_support_var = IntVar()
    regional_support_var = IntVar()
    office_admin_practice_var = IntVar()
    office_admin_provider_var = IntVar()
    provider_var = IntVar()
    cozeva_support_checkbox = Checkbutton(input_frame_2, text="Cozeva Support", variable=cozeva_support_var, font=("Nunito Sans", 10))
    limited_cozeva_support_checkbox = Checkbutton(input_frame_2, text="Limited Cozeva Support", variable=limited_cozeva_support_var, font=("Nunito Sans", 10))
    customer_support_checkbox = Checkbutton(input_frame_2, text="Customer Support", variable=customer_support_var, font=("Nunito Sans", 10))
    regional_support_checkbox = Checkbutton(input_frame_2, text="Regional support", variable=regional_support_var, font=("Nunito Sans", 10))
    office_admin_practice_checkbox = Checkbutton(input_frame_2, text="Office Admin(Practice Delegate)", variable=office_admin_practice_var, font=("Nunito Sans", 10))
    office_admin_provider_checkbox = Checkbutton(input_frame_2, text="Office Admin(Provider Delegate)", variable=office_admin_provider_var, font=("Nunito Sans", 10))
    provider_checkbox = Checkbutton(input_frame_2, text="Provider", variable=provider_var, font=("Nunito Sans", 10))
    # customer_support_username_var = StringVar()
    # regional_support_username_var = StringVar()
    # office_admin_practice_username_var = StringVar()
    # office_admin_provider_username_var = StringVar()
    # provider_username_var = StringVar()
    limited_cozeva_support_textbox = Entry(input_frame_2)
    customer_support_textbox = Entry(input_frame_2)
    regional_support_textbox = Entry(input_frame_2)
    office_admin_practice_textbox = Entry(input_frame_2)
    office_admin_provider_textbox = Entry(input_frame_2)
    provider_textbox = Entry(input_frame_2)
    nextbutton2 = Button(input_frame_2, text="Lock choices and Proceed", command=frame2next, font=("Nunito Sans", 10))
    prevbutton2 = Button(input_frame_2, text="Go back", command=frame2prev, font=("Nunito Sans", 10))

    #widgets for frame3
    #this frame will contain a series of checkboxes that will select verification scope
    frame3_info_label = Label(input_frame_3, text="Select Automated Test Scope", font=("Nunito Sans", 10))
    support_sidemenu_var = IntVar()
    practice_sidemenu_var = IntVar()
    provider_sidemenu_var = IntVar()
    patient_dashboard_var = IntVar()
    provider_registry_var = IntVar()
    practice_registry_var = IntVar()
    support_registry_var = IntVar()
    global_search_var = IntVar()
    provider_mspl_var = IntVar()
    time_capsule_var = IntVar()
    secure_messaging_var = IntVar()
    accordion_verification_var = IntVar()
    hcc_validation_var = IntVar()
    select_var = IntVar()
    select_checkbox = Checkbutton(input_frame_3, text="Select All", variable=select_var, command=select_all, font=("Nunito Sans", 10))
    support_sidemenu_checkbox = Checkbutton(input_frame_3, text="Support Sidemenu", variable=support_sidemenu_var, font=("Nunito Sans", 10))
    practice_sidemenu_checkbox = Checkbutton(input_frame_3, text="Practice Sidemenu", variable=practice_sidemenu_var, font=("Nunito Sans", 10))
    provider_sidemenu_checkbox = Checkbutton(input_frame_3, text="provider Sidemenu", variable=provider_sidemenu_var, font=("Nunito Sans", 10))
    patient_dashboard_checkbox = Checkbutton(input_frame_3, text="Patient Dashboard", variable=patient_dashboard_var, font=("Nunito Sans", 10))
    provider_registry_checkbox = Checkbutton(input_frame_3, text="Provider Registry", variable=provider_registry_var, font=("Nunito Sans", 10))
    practice_registry_checkbox = Checkbutton(input_frame_3, text="Practice Registry", variable=practice_registry_var, font=("Nunito Sans", 10))
    support_registry_checkbox = Checkbutton(input_frame_3, text="Support Registry", variable=support_registry_var, font=("Nunito Sans", 10))
    global_search_checkbox = Checkbutton(input_frame_3, text="Global Search", variable=global_search_var, font=("Nunito Sans", 10))
    provider_mspl_checkbox = Checkbutton(input_frame_3, text="Provider's MSPL", variable=provider_mspl_var, font=("Nunito Sans", 10))
    time_capsule = Checkbutton(input_frame_3, text="Time Capsule", variable=time_capsule_var, font=("Nunito Sans", 10))
    secure_messaging_checkbox = Checkbutton(input_frame_3, text="Secure Messaging", variable=secure_messaging_var, font=("Nunito Sans", 10))
    accordion_verification_checkbox = Checkbutton(input_frame_3, text="Accordion and counts Verification", variable=accordion_verification_var,
                                            font=("Nunito Sans", 10))
    hcc_validation_checkbox = Checkbutton(input_frame_3, text="HCC Validation", variable=hcc_validation_var, font=("Nunito Sans", 10))
    nextbutton3 = Button(input_frame_3, text="Start Automated Test", command=frame3next, font=("Nunito Sans", 10))
    prevbutton3 = Button(input_frame_3, text="Go Back", command=frame3prev, font=("Nunito Sans", 10))

    #widgets for new_launch_frame
    new_launch_frame_label = Label(new_launch_frame, text="Select Scope of Test", font=("Nunito Sans", 10))
    all_navigation_var = IntVar()
    all_navigation_checkbox = Checkbutton(new_launch_frame, text="Validate Multi-Context Navigation", variable=all_navigation_var, font=("Nunito Sans", 10))
    LoB_Measure_var = IntVar()
    LoB_Measure_checkbox =  Checkbutton(new_launch_frame, text="Validate Measures LoB-wise", variable=LoB_Measure_var, font=("Nunito Sans", 10))
    provider_tab_var = IntVar()
    provider_tab_checkbox = Checkbutton(new_launch_frame, text="Validate Metric Specific Provider List", variable=provider_tab_var, font=("Nunito Sans", 10))
    market_sheet_var = IntVar()
    market_sheet_checkbox = Checkbutton(new_launch_frame, text="Validate Market Sheet Contents", variable=market_sheet_var, font=("Nunito Sans", 10))
    patient_medication_var = IntVar()
    patient_medication_checkbox = Checkbutton(new_launch_frame, text="Validate Patient Medication", variable=patient_medication_var, font=("Nunito Sans", 10))
    MaP_var = IntVar()
    MaP_checkbox = Checkbutton(new_launch_frame, text="Validate Mark as Pending", variable=MaP_var, font=("Nunito Sans", 10))
    apptray_var = IntVar()
    apptray_checkbox = Checkbutton(new_launch_frame, text="Validate Apptray Access", variable=apptray_var, font=("Nunito Sans", 10))
    train_resource_var = IntVar()
    train_resource_checkbox = Checkbutton(new_launch_frame, text="Validate Training/Resources Page", variable=train_resource_var, font=("Nunito Sans", 10))
    userlist_var = IntVar()
    userlist_checkbox = Checkbutton(new_launch_frame, text="Validate Userlist and User Creation", variable=userlist_var, font=("Nunito Sans", 10))
    banner_announce_var = IntVar()
    banner_announce_checkbox = Checkbutton(new_launch_frame, text="Validate Banners/Announcements", variable=banner_announce_var, font=("Nunito Sans", 10))
    ce_toggle_var = IntVar()
    ce_toggle_checkbox = Checkbutton(new_launch_frame, text="Validate CE Toggle States", variable=ce_toggle_var, font=("Nunito Sans", 10))
    NL_accordion_verification_var = IntVar()
    NL_accordion_verification_checkbox = Checkbutton(new_launch_frame, text="Validate Support Level Accordion Measures", variable=NL_accordion_verification_var, font=("Nunito Sans", 10))
    NL_global_search_var = IntVar()
    NL_global_search_checkbox = Checkbutton(new_launch_frame, text="Validate Global Search", variable=NL_global_search_var, font=("Nunito Sans", 10))
    sticket_var = IntVar()
    sticket_checkbox = Checkbutton(new_launch_frame, text="Validate Sticket Logs", variable=sticket_var, font=("Nunito Sans", 10))
    NL_patient_dashboard_var = IntVar()
    NL_patient_dashboard_checkbox = Checkbutton(new_launch_frame, text="Validate Patient Dashboard and Timeline", variable=NL_patient_dashboard_var, font=("Nunito Sans", 10))
    coding_tool_var = IntVar()
    coding_tool_checkbox = Checkbutton(new_launch_frame, text="Validate Coding tool", variable=coding_tool_var, font=("Nunito Sans", 10))
    denom_eligibility_var = IntVar()
    denom_eligibility_checkbox = Checkbutton(new_launch_frame, text="Validate Denominator Eligibility for CBP/HbA1c Measures", variable=denom_eligibility_var, font=("Nunito Sans", 10))
    NL_hcc_validation_var = IntVar()
    NL_hcc_validation_checkbox = Checkbutton(new_launch_frame, text="HCC Validation", variable=NL_hcc_validation_var, font=("Nunito Sans", 10))
    NL_select_var = IntVar()
    NL_select_checkbox = Checkbutton(new_launch_frame, text="Select All", variable=NL_select_var, command=NL_select_all, font=("Nunito Sans", 10))
    NL_next_button = Button(new_launch_frame, text="Begin Test", command=nlnext, font=("Nunito Sans", 10))
    NL_custID = Entry(new_launch_frame, text="Customer ID")

    def buildExistingSchema():
        global grid_row
        current_schema = sp.getCurrentSchema()
        for schema_row in current_schema:
            Cust_ID_list.append(Entry(Analytics_frame))
            Measurement_year.append(Entry(Analytics_frame))
            medicare.append(Entry(Analytics_frame))
            commercial.append(Entry(Analytics_frame))
            utilization.append(Entry(Analytics_frame))
            usage.append(Entry(Analytics_frame))
            Cust_ID_list[grid_row - 2].grid(row=grid_row, column=0)
            Measurement_year[grid_row - 2].grid(row=grid_row, column=1)
            medicare[grid_row - 2].grid(row=grid_row, column=2)
            commercial[grid_row - 2].grid(row=grid_row, column=3)
            utilization[grid_row - 2].grid(row=grid_row, column=4)
            usage[grid_row - 2].grid(row=grid_row, column=5)
            Cust_ID_list[grid_row - 2].insert(0, schema_row[0])
            Measurement_year[grid_row - 2].insert(0, schema_row[2])
            medicare[grid_row - 2].insert(0, schema_row[3])
            commercial[grid_row - 2].insert(0, schema_row[4])
            utilization[grid_row - 2].insert(0, schema_row[5])
            usage[grid_row - 2].insert(0, schema_row[6])
            grid_row+=1
    def new_schema():
        global grid_row
        Cust_ID_list.append(Entry(Analytics_frame))
        Measurement_year.append(Entry(Analytics_frame))
        medicare.append(Entry(Analytics_frame))
        commercial.append(Entry(Analytics_frame))
        utilization.append(Entry(Analytics_frame))
        usage.append(Entry(Analytics_frame))
        Cust_ID_list[grid_row-2].grid(row=grid_row, column=0)
        Measurement_year[grid_row - 2].grid(row=grid_row, column=1)
        medicare[grid_row - 2].grid(row=grid_row, column=2)
        commercial[grid_row - 2].grid(row=grid_row, column=3)
        utilization[grid_row - 2].grid(row=grid_row, column=4)
        usage[grid_row - 2].grid(row=grid_row, column=5)
        #print(grid_row)
        #print(len(Cust_ID_list))
        grid_row = grid_row + 1
    def getSchema():
        dda = []
        temp = []
        temp.clear()
        schema_list_length = len(Cust_ID_list)
        for i in range(schema_list_length):
            #temp.clear()
            temp = []
            temp.append(Cust_ID_list[i].get())
            temp.append(db.fetchCustomerName(Cust_ID_list[i].get()))
            temp.append(Measurement_year[i].get())
            temp.append(medicare[i].get())
            temp.append(commercial[i].get())
            temp.append(utilization[i].get())
            temp.append(usage[i].get())
            print(dda)
            dda.append(temp)
            print(dda)
            print(temp)
            print(i)
            #temp.clear()
            #print(dda)
        print(dda)
        sp.loadSchema(dda)
        root.destroy()
    #widgets for Analytics_frame
    Schema_info_label = Label(Analytics_frame, text="Edit/Add a new verification Schema")
    add_new_entry = Button(Analytics_frame, text='+', command=new_schema)
    #example_label = Label(Analytics_frame, text="Example: '200,2022,1500,2021,4600,2022,3000,2021")
    #customer_choice_entry = Entry(Analytics_frame)
    Cust_ID_list_header = Label(Analytics_frame, text="Customer ID")
    Measurement_year_header = Label(Analytics_frame, text="Measurement Year")
    medicare_header = Label(Analytics_frame, text="Medicare")
    commercial_header = Label(Analytics_frame, text="Commercial")
    utilization_header = Label(Analytics_frame, text="Utilization")
    usage_header = Label(Analytics_frame, text="Usage")
    Cust_ID_list = []
    Measurement_year = []
    medicare = []
    commercial = []
    utilization = []
    usage = []
    analytics_next= Button(Analytics_frame, text="Lock choices and Proceed", command=getSchema)
    #packing elements into Analytics_frame
    Schema_info_label.grid(row=0, columnspan=4)
    #example_label.grid(row=1, columnspan=5)
    #customer_choice_entry.grid(row=2, columnspan=5)
    add_new_entry.grid(row=0,column=5)
    analytics_next.grid(row=0, column=6)
    Cust_ID_list_header.grid(row=1, column=0)
    Measurement_year_header.grid(row=1, column=1)
    medicare_header.grid(row=1, column=2)
    commercial_header.grid(row=1, column=3)
    utilization_header.grid(row=1, column=4)
    usage_header.grid(row=1, column=5)



    #packing elements into frame 1
    customer_label.grid(row=0, column=0, columnspan=5)
    customer_drop.grid(row=1, column=0, columnspan=5)
    context_label.grid(row=2, column=2)
    config_label.grid(row=2, column=4)
    config_cozeva.grid(row=3, column=2, sticky='w')
    Checkbox_cozeva.grid(row=4, column=2, sticky="e")
    radiobutton_default.grid(row=3, column=4, sticky="w")
    Checkbox_analytics.grid(row=5, column=2, sticky="e")
    config_analytics.grid(row=6, column=2, sticky='w')
    radiobutton_illchoose.grid(row=4, column=4, sticky="w")
    radiobutton_newcustomer.grid(row=5, column=4, sticky="w")
    window_location_label.grid(row=7, column=0, columnspan=5)
    radiobutton_window_left.grid(row=8, column=2, sticky="e")
    radiobutton_window_right.grid(row=8, column=3, sticky="w")
    radiobutton_env_label.grid(row=9, column=2)
    radiobutton_env_stage.grid(row=10, column=2, sticky="w")
    radiobutton_headless_label.grid(row=9, column=4)
    radiobutton_headless_yes.grid(row=10, column=4, sticky="w")
    radiobutton_headless_no.grid(row=11, column=4, sticky="w")
    nextbutton1.grid(row=13, column=0, columnspan=5, pady=45)
    Checkbox_cozeva.select()


    #packing elements into frame 2

    frame2_info_label.grid(row=0, column=0, columnspan=5)
    cozeva_support_checkbox.grid(row=1, column=0, columnspan=5, sticky="w")
    limited_cozeva_support_checkbox.grid(row=2, column=0, columnspan=5, sticky="w")
    limited_cozeva_support_textbox.grid(row=3, column=0, columnspan=5, sticky="w")
    customer_support_checkbox.grid(row=4, column=0, columnspan=5, sticky="w")
    customer_support_textbox.grid(row=5, column=0, columnspan=5, sticky="w")
    regional_support_checkbox.grid(row=6, column=0, columnspan=5, sticky="w")
    regional_support_textbox.grid(row=7, column=0, columnspan=5, sticky="w")
    office_admin_practice_checkbox.grid(row=8, column=0, columnspan=5, sticky="w")
    office_admin_practice_textbox.grid(row=9, column=0, columnspan=5, sticky="w")
    office_admin_provider_checkbox.grid(row=10, column=0, columnspan=5, sticky="w")
    office_admin_provider_textbox.grid(row=11, column=0, columnspan=5, sticky="w")
    provider_checkbox.grid(row=12, column=0, columnspan=5, sticky="w")
    provider_textbox.grid(row=13, column=0, columnspan=5, sticky="w")
    nextbutton2.grid(row=14, column=3)
    prevbutton2.grid(row=14, column=0)

    #packing widgets into frame 3
    frame3_info_label.grid(row=0, column=0, columnspan=5)
    select_checkbox.grid(row=1, column=0, columnspan=5, sticky="w")
    support_sidemenu_checkbox.grid(row=2, column=0, columnspan=5, sticky="w")
    practice_sidemenu_checkbox.grid(row=3, column=0, columnspan=5, sticky="w")
    provider_sidemenu_checkbox.grid(row=4, column=0, columnspan=5, sticky="w")
    patient_dashboard_checkbox.grid(row=5, column=0, columnspan=5, sticky="w")
    provider_registry_checkbox.grid(row=6, column=0, columnspan=5, sticky="w")
    practice_registry_checkbox.grid(row=7, column=0, columnspan=5, sticky="w")
    support_registry_checkbox.grid(row=8, column=0, columnspan=5, sticky="w")
    global_search_checkbox.grid(row=9, column=0, columnspan=5, sticky="w")
    provider_mspl_checkbox.grid(row=10, column=0, columnspan=5, sticky="w")
    time_capsule.grid(row=11, column=0, columnspan=5, sticky="w")
    secure_messaging_checkbox.grid(row=12, column=0, columnspan=5, sticky="w")
    accordion_verification_checkbox.grid(row=13, column=0, columnspan=5, sticky="w")
    hcc_validation_checkbox.grid(row=14, column=0, columnspan=5, sticky="w")
    nextbutton3.grid(row=15, column=3)
    prevbutton3.grid(row=15, column=0)

    # pack elements into new_launch_frame
    new_launch_frame_label.grid(row=0, column=0, columnspan=5)
    NL_select_checkbox.grid(row=1, column=0, columnspan=5, sticky="w")
    all_navigation_checkbox.grid(row=2, column=0, columnspan=5, sticky="w")
    LoB_Measure_checkbox.grid(row=3, column=0, columnspan=5, sticky="w")
    LoB_Measure_checkbox.config(state='disabled')
    provider_tab_checkbox.grid(row=4, column=0, columnspan=5, sticky="w")
    market_sheet_checkbox.grid(row=5, column=0, columnspan=5, sticky="w")
    patient_medication_checkbox.grid(row=6, column=0, columnspan=5, sticky="w")
    MaP_checkbox.grid(row=7, column=0, columnspan=5, sticky="w")
    apptray_checkbox.grid(row=8, column=0, columnspan=5, sticky="w")
    train_resource_checkbox.grid(row=9, column=0, columnspan=5, sticky="w")
    userlist_checkbox.grid(row=10, column=0, columnspan=5, sticky="w")
    userlist_checkbox.config(state='disabled')
    banner_announce_checkbox.grid(row=11, column=0, columnspan=5, sticky="w")
    banner_announce_checkbox.config(state='disabled')
    ce_toggle_checkbox.grid(row=12, column=0, columnspan=5, sticky="w")
    NL_accordion_verification_checkbox.grid(row=13, column=0, columnspan=5, sticky="w")
    NL_global_search_checkbox.grid(row=14, column=0, columnspan=5, sticky="w")
    sticket_checkbox.grid(row=15, column=0, columnspan=5, sticky="w")
    NL_patient_dashboard_checkbox.grid(row=16, column=0, columnspan=5, sticky="w")
    #coding_tool_checkbox.grid(row=17, column=0, columnspan=5, sticky="w")
    denom_eligibility_checkbox.grid(row=17, column=0, columnspan=5, sticky="w")
    denom_eligibility_checkbox.config(state='disabled')
    NL_hcc_validation_checkbox.grid(row=18, column=0, columnspan=5, sticky="w")
    NL_custID.grid(row=19, column=0, columnspan=5)
    NL_next_button.grid(row=20, column=0, columnspan=5)


    #packing frame 1 into root
    input_frame_1.pack()

    try:

        print("Test 1")
        options = webdriver.ChromeOptions()
        #options.add_argument("--disable-notifications")
        #options.add_argument("--start-maximized")
        options.add_argument(locator.chrome_profile_path)  # Path to your chrome profile
        # options.add_argument("--headless")
        # options.add_argument('--disable-gpu')
        # options.add_argument("--window-size=1920,1080")
        # options.add_argument("--start-maximized")
        driver = webdriver.Chrome(executable_path=locator.chrome_driver_path, options=options)
        time.sleep(1)
        driver.quit()

    except Exception as e:
        warninglabel.grid(row=13, columnspan=5, sticky="w")
        traceback.print_exc()
        time.sleep(1)






    root.title("Cozeva Stage Verification")
    root.iconbitmap("assets/icon.ico")
    #root.geometry("400x400")
    root.mainloop()



# launchgui()
# print(verification_specs)


