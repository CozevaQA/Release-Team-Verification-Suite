import tkinter as tk
import webbrowser
import os
import ExcelProcessor as db


def open_html_file(file_path):
    file_url = f"file:///{os.path.abspath(file_path)}"
    webbrowser.open_new_tab(file_url)


def open_xl_file(filename):
    os.startfile(os.path.join(os.getcwd(), 'assets/SchemaForAllWorksheets.xlsx'))


def create_ui(directory_path):
    root = tk.Tk()
    root.title("Chart List Export Reports")
    root.iconbitmap("assets/icon.ico")

    # Create frames for two columns
    frame_obs = tk.Frame(root)
    frame_reports = tk.Frame(root)

    frame_obs.grid(row=0, column=0, padx=10, pady=10)
    frame_reports.grid(row=0, column=1, padx=10, pady=10)

    # Add labels for columns
    tk.Label(frame_obs, text="Observations").pack()
    tk.Label(frame_reports, text="Reports").pack()

    html_files = [f for f in os.listdir(directory_path) if f.endswith('.html')]

    if not html_files:
        label = tk.Label(root, text="No HTML files found in the specified directory.")
        label.pack(pady=10)
    else:
        for html_file in html_files:
            file_path = os.path.join(directory_path, html_file)

            if "Observations" in html_file:
                client_name = db.fetchCustomerName(html_file.replace("_Observations.html", ""))
                html_file = client_name
                button = tk.Button(frame_obs, text=html_file, command=lambda fp=file_path: open_html_file(fp))
                button.pack(pady=5)
            elif "Reports" in html_file:
                client_name = db.fetchCustomerName(html_file.replace("_Reports.html", ""))
                html_file = client_name
                button = tk.Button(frame_reports, text=html_file, command=lambda fp=file_path: open_html_file(fp))
                button.pack(pady=5)

    root.mainloop()


directory_path = r'C:\VerificationReports\Chart List Reports'
create_ui(directory_path)
