from tkinter import *
from PIL import ImageTk, Image
from tkinter import ttk

def master_gui():
    root = Tk()
    root.configure(background='white')
    style = ttk.Style()
    style.theme_use('alt')
    style.configure('My.TButton', font=('Helvetica', 13, 'bold'), foreground='Black', background='#5a9c32', padding=15, highlightthickness=0, height=1, width=25)

    #style.configure('My.TButton', font=('American typewriter', 14), background='#232323', foreground='white')
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
    multi_role_image = ImageTk.PhotoImage(image_sizer("assets/images/Multi_role_access.png"))
    special_column_image = ImageTk.PhotoImage(image_sizer("assets/images/special_columns.png"))
    hospital_activity_image = ImageTk.PhotoImage(image_sizer("assets/images/hospital_activity.png"))
    supp_data_image = ImageTk.PhotoImage(image_sizer("assets/images/supp_data.png"))
    cozeva_logo_image = ImageTk.PhotoImage(Image.open("assets/images/cozeva_logo.png").resize((300, 66)))

    #Widgets+

    logo_label = Label(root, image=cozeva_logo_image, background="white")
    logo_label.grid(row=0, column=1)

    root.columnconfigure(1, weight=2)
    root.rowconfigure(0, weight=2)
    logo_label.grid(sticky="w")
    please_select_label = Label(root, text="Release Team Verification Suite", background="white", font=("Times New Roman", 18))
    please_select_label.grid(row=1, column=1, columnspan=3, sticky='w')
    #TRYING SOMETHING ELSE, HOPING THIS WORKS ITS 3 AM

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
        #import Supplemental_data_alternate

    def on_conf_dis():
        root.destroy()






    button_widgets = []
    button_widgets.append(
        ttk.Button(root, text="First time Setup", command=on_first_time_setup, image=first_time_setup_image,
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
        ttk.Button(root, text="Multi-role Access Check", command=on_role_access, image=multi_role_image,
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
        ttk.Button(root, text="Confirm/Disconfirm(WIP)", command=on_conf_dis, image=hcc_validation_image,
                   compound="left", style='My.TButton'))

    widget_counter = 0
    loopbreak = 0
    for i in range(2, 6):
        for j in range(3):
            try:
                button_widgets[widget_counter].grid(row=i, column=j, padx=5, pady=5)
            except IndexError as e:
                loopbreak=1
                break
            widget_counter += 1
        if loopbreak == 1:
            break

    root.title("Release Team Master Suite")
    root.iconbitmap("assets/icon.ico")
    #root.geometry("400x400+300+100")
    root.mainloop()


master_gui()
