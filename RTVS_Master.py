from tkinter import *
from PIL import ImageTk, Image

def master_gui():
    root = Tk()

    def on_nextbutton():
        x=0

    #Widgets
    please_select_label = Label(root, text="Select Script", font=("Nunito Sans", 10))
    # script_var = IntVar()
    # daily_validation_radio = Radiobutton(root, text="Verification Suite", variable=script_var, value=0, font=("Nunito Sans", 10))
    # hcc_validation_radio = Radiobutton(root, text="Verification Suite", variable=script_var, value=1,
    #                                      font=("Nunito Sans", 10))
    # global_search_radio = Radiobutton(root, text="Verification Suite", variable=script_var, value=2,
    #                                      font=("Nunito Sans", 10))
    # next_button = Button(root, text="Next", command=on_nextbutton, font=("Nunito Sans", 10))
    #
    # #load widgets onto root
    please_select_label.grid(row=0, column=1, columnspan=5, sticky="w")
    # daily_validation_radio.grid(row=1, column=1, columnspan=5, sticky="w")
    # hcc_validation_radio.grid(row=2, column=1, columnspan=5, sticky="w")
    # global_search_radio.grid(row=3, column=1, columnspan=5, sticky="w")
    # next_button.grid(row=4, column=1, columnspan=5)

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

    def image_sizer(image_path):
        image_small = Image.open(image_path).resize((25, 25))

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

    Button(root, text="First time Setup", command=on_first_time_setup, font=("Nunito Sans", 10)).grid(row=1,
                                                                                                      column=1,
                                                                                                      columnspan=4,
                                                                                                      sticky="w")
    Button(root, text="Verification Suite", command=on_verification_suite, font=("Nunito Sans", 10)).grid(row=2, column=1, columnspan=4, sticky="w")
    Button(root, text="HCC Validation Multi (Custom client list)", command=on_hcc_validation, font=("Nunito Sans", 10)).grid(row=3,
                                                                                                        column=1,
                                                                                                        columnspan=4,
                                                                                                        sticky="w")
    Button(root, text="Global Search", command=on_global_search, font=("Nunito Sans", 10)).grid(row=4,
                                                                                                column=1,
                                                                                                columnspan=4,
                                                                                                sticky="w")
    Button(root, text="Task Ingestion (AWV Only, Others WIP)", command=on_task_ingestion, font=("Nunito Sans", 10)).grid(row=5,
                                                                                                                        column=1,
                                                                                                                        columnspan=4,
                                                                                                                        sticky="w")
    Button(root, text="Analytics (Default Config Schema)", command=on_analytics, font=("Nunito Sans", 10)).grid(row=6,
                                                                                                                column=1,
                                                                                                                columnspan=4,
                                                                                                                sticky="w")
    Button(root, text="Slow Log Trends", command=on_slow_trends, font=("Nunito Sans", 10)).grid(row=7,
                                                                                                column=1,
                                                                                                columnspan=4,
                                                                                                sticky="w")
    Button(root, text="Multi-role Access Check", command=on_role_access, font=("Nunito Sans", 10)).grid(row=8, column=1, columnspan=5, sticky="w")
    Button(root, text="Special Columns", command=on_special_columns, font=("Nunito Sans", 10)).grid(row=9,
                                                                                                column=1,
                                                                                                columnspan=4,
                                                                                                sticky="w")
    Button(root, text="Hospital Activity (All Clients)", command=on_hospital_activity, font=("Nunito Sans", 10)).grid(row=10,
                                                                                                    column=1,
                                                                                                    columnspan=4,
                                                                                                    sticky="w")
    Label(root, image=first_time_setup_image, width=40, height=40).grid(row=1, column=0, sticky="w")
    Label(root, image=verification_suite_image, width=40, height=40).grid(row=2, column=0, sticky="w")
    Label(root, image=hcc_validation_image, width=40, height=40).grid(row=3, column=0, sticky="w")
    Label(root, image=global_search_image, width=40, height=40).grid(row=4, column=0, sticky="w")
    Label(root, image=task_ingestion_image, width=40, height=40).grid(row=5, column=0, sticky="w")
    Label(root, image=analytics_image, width=40, height=40).grid(row=6, column=0, sticky="w")
    Label(root, image=slow_log_image, width=40, height=40).grid(row=7, column=0, sticky="w")
    Label(root, image=multi_role_image, width=40, height=40).grid(row=8, column=0, sticky="w")
    Label(root, image=special_column_image, width=40, height=40).grid(row=9, column=0, sticky="w")
    Label(root, image=hospital_activity_image, width=40, height=40).grid(row=10, column=0, sticky="w")








    root.title("Release Team Master Suite")
    root.iconbitmap("assets/icon.ico")
    root.mainloop()


master_gui()
