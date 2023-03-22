from tkinter import *
from PIL import ImageTk, Image

def master_gui():
    root = Tk()
    root.configure(background='white')

    def on_nextbutton():
        x=0

    #Widgets
    please_select_label = Label(root, text="Select Script", background="white", font=("Times New Roman", 20))
    please_select_label.grid(row=0, column=1, columnspan=5)
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
        import Supplemental_data_alternate

    def on_conf_dis():
        root.destroy()


    def image_sizer(image_path):
        image_small = Image.open(image_path).resize((100, 100))

        return image_small
    #making image widgets
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

    Button(root, text="First time Setup", command=on_first_time_setup, image=first_time_setup_image, compound="top", font=("Nunito Sans", 10)).grid(row=1, column=0, padx=40, pady=20)
    Button(root, text="Verification Suite", command=on_verification_suite, image=first_time_setup_image, compound="top", font=("Nunito Sans", 10)).grid(row=1, column=1, padx=40, pady=20)
    Button(root, text="HCC Validation Multi (Custom client list)", command=on_hcc_validation, image=first_time_setup_image, compound="top", font=("Nunito Sans", 10)).grid(row=1, column=2, padx=40, pady=20)
    Button(root, text="Global Search", command=on_global_search, image=first_time_setup_image, compound="top", font=("Nunito Sans", 10)).grid(row=4, column=1, sticky="w", padx=40, pady=20)
    Button(root, text="Task Ingestion (AWV Only, Others WIP)", command=on_task_ingestion, image=first_time_setup_image, compound="top", font=("Nunito Sans", 10)).grid(row=5, column=1, sticky="w", padx=40, pady=20)
    Button(root, text="Analytics (Default Config Schema)", command=on_analytics, image=first_time_setup_image, compound="top", font=("Nunito Sans", 10)).grid(row=6, column=1, sticky="w", padx=40, pady=20)
    Button(root, text="Slow Log Trends", command=on_slow_trends, image=first_time_setup_image, compound="top", font=("Nunito Sans", 10)).grid(row=7, column=1, sticky="w", padx=40, pady=20)
    Button(root, text="Multi-role Access Check", command=on_role_access, image=first_time_setup_image, compound="top", font=("Nunito Sans", 10)).grid(row=8, column=1, sticky="w", padx=40, pady=20)
    Button(root, text="Special Columns", command=on_special_columns, image=first_time_setup_image, compound="top", font=("Nunito Sans", 10)).grid(row=9, column=1, sticky="w", padx=40, pady=20)
    Button(root, text="Hospital Activity (All Clients)", command=on_hospital_activity, image=first_time_setup_image, compound="top", font=("Nunito Sans", 10)).grid(row=10, column=1, sticky="w", padx=40, pady=20)
    Button(root, text="Supplemental Data Addition", command=on_supp_data, image=first_time_setup_image, compound="top", font=("Nunito Sans", 10)).grid(row=11, column=1, sticky="w", padx=40, pady=20)
    Button(root, text="Confirm/Disconfirm(WIP)", command=on_conf_dis, image=first_time_setup_image, compound="top", font=("Nunito Sans", 10)).grid(row=12, column=1, sticky="w", padx=40, pady=20)









    root.title("Release Team Master Suite")
    root.iconbitmap("assets/icon.ico")
    root.mainloop()


master_gui()
