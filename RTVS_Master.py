from tkinter import *

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



    Button(root, text="Verification Suite", command=on_verification_suite, font=("Nunito Sans", 10)).grid(row=1, column=1, columnspan=5, sticky="w")
    Button(root, text="HCC Validation Multi", command=on_hcc_validation, font=("Nunito Sans", 10)).grid(row=2,
                                                                                                        column=1,
                                                                                                        columnspan=5,
                                                                                                        sticky="w")
    Button(root, text="Global Search", command=on_global_search, font=("Nunito Sans", 10)).grid(row=3,
                                                                                                column=1,
                                                                                                columnspan=5,
                                                                                                sticky="w")
    Button(root, text="Task Ingestion (AWV Only, Others WIP)", command=on_task_ingestion, font=("Nunito Sans", 10)).grid(row=4,
                                                                                                                        column=1,
                                                                                                                        columnspan=5,
                                                                                                                        sticky="w")
    Button(root, text="Analytics (Default Config Schema)", command=on_analytics, font=("Nunito Sans", 10)).grid(row=5,
                                                                                                                column=1,
                                                                                                                columnspan=5,
                                                                                                                sticky="w")
    Button(root, text="Slow Log Trends", command=on_slow_trends, font=("Nunito Sans", 10)).grid(row=6,
                                                                                                column=1,
                                                                                                columnspan=5,
                                                                                                sticky="w")
    Button(root, text="Multi-role Access Check", command=on_role_access, font=("Nunito Sans", 10)).grid(row=7, column=1, columnspan=5, sticky="w")



    root.title("Release Team Master Suite")
    root.iconbitmap("assets/icon.ico")
    root.mainloop()


master_gui()
