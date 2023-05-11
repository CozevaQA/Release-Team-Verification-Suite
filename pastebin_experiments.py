import requests
import tkinter as tk
import xml.etree.ElementTree as ET
import ExcelProcessor as db
from PIL import ImageTk, Image


# Set the API key and the unique ID of the paste you want to read
api_key = "60AyzJ0aLaF8ryKbw84fnB1cTnO7qjoI"
paste_key = ""
username = "wdey"
password = "potterfan@85"
api_user_key = ""



def fetch_paste():
    # Set the URL of the API endpoint for reading pastes
    paste_url = f"https://pastebin.com/raw/{paste_key}"

    # Make a GET request to the API endpoint to fetch the content of the paste
    fetch_response = requests.get(paste_url)

    # Check if the request was successful
    if fetch_response.status_code == 200:
        # Print the content of the paste
        return fetch_response.text
    else:
        return "No data"


def update_paste(data):
    return data


#print(fetch_paste())
def fetch_userkey():
    # Set the parameters for logging in and getting the user key
    params = {
        "api_user_name": username,
        "api_user_password": password,
        "api_dev_key": api_key
    }

    # Make a POST request to the Pastebin API to log in and get the user key
    response_user = requests.post("https://pastebin.com/api/api_login.php", data=params)

    # Check if the request was successful
    if response_user.status_code == 200:
        # Get the user key
        global api_user_key
        api_user_key = response_user.text.strip()

        print(f"api_user_key: {api_user_key}")
    else:
        print("Error logging in and getting user key.")
        exit()


def modify_paste(paste_content):
    # Modify the content of the paste
    modified_content = paste_content.replace("55", "555555")


    # Set the URL of the API endpoint for modifying pastes
    modify_url = "https://pastebin.com/api/api_post.php"

    # Set the parameters for modifying the paste
    params = {
        "api_option": "paste",
        "api_user_key": api_user_key,
        "api_dev_key": api_key,
        "api_paste_private": 0,
        "api_paste_code": modified_content
    }

    # Make a POST request to the API endpoint to modify the paste
    response = requests.post(modify_url, data=params)

    # Check if the request was successful
    if response.status_code == 200:
        print("Paste added successfully.")
    else:
        print("Paste addition failed.")


def fetch_pastekey():
    list_url = "https://pastebin.com/api/api_post.php"
    params = {
        "api_option": "list",
        "api_user_key": api_user_key,
        "api_dev_key": api_key
    }
    response = requests.post(list_url, data=params)
    if response.status_code == 200:
        paste_root = ET.fromstring(response.text)
        global paste_key
        paste_key = paste_root.find("paste_key").text
    else:
        print("Error retrieving list of pastes.")
        exit()


def delete_paste():
    delete_url = "https://pastebin.com/api/api_post.php"
    params = {
        "api_option": "delete",
        "api_user_key": api_user_key,
        "api_dev_key": api_key,
        "api_paste_key": paste_key
    }
    response = requests.post(delete_url, data=params)
    if response.status_code == 200:
        print("Paste deleted successfully.")
    else:
        print("Paste deletion failed.")

def gui(data):
    data_list = [line.split() for line in data.splitlines()]

    # Create the main application window
    root = tk.Tk()

    status_list = ["Not Started", "Ongoing", "Completed"]
    green_dot_image = ImageTk.PhotoImage(Image.open("assets/images/GreenDot.png").resize((10, 10)))
    red_dot_image = ImageTk.PhotoImage(Image.open("assets/images/RedDot.png").resize((10, 10)))
    orange_dot_image = ImageTk.PhotoImage(Image.open("assets/images/OrangeDot.png").resize((10, 10)))

    image_list = [red_dot_image, orange_dot_image, green_dot_image]

    # Display the data in a grid
    for client_row, row in enumerate(data_list):
        for grid_pointer, grid_data in enumerate(row):
            if grid_pointer == 1:
                label = tk.Label(root, text=status_list[int(grid_data)], image=image_list[int(grid_data)], compound='left', padx=5, pady=5)
                label.grid(row=client_row, column=grid_pointer, sticky="w")
                count_button = tk.Button(root, text="Run Count Validation", state="disabled", bg="red")
                count_button.grid(row=client_row, column=3, sticky="news")
                if int(grid_data) == 2:
                    count_button.config(state='active', bg="green")
            elif grid_pointer == 0:
                label = tk.Label(root, text=db.fetchCustomerName(str(grid_data)), padx=5, pady=5)
                label.grid(row=client_row, column=grid_pointer, sticky="w")
            else:
                label = tk.Label(root, text=grid_data, padx=5, pady=5)
                label.grid(row=client_row, column=grid_pointer, sticky="w")

    # Run the Tkinter event loop
    root.title("Computation Status Checker")
    root.iconbitmap("assets/icon.ico")
    # root.geometry("400x400+300+100")
    root.mainloop()



# Retrieve paste data

# fetch_userkey()
# fetch_pastekey()
# print(paste_key)
# paste_content = fetch_paste()
# print(paste_content)
#
# # Update the local paste content with latest data source
# paste_content = update_paste(paste_content)
# # update it in pastebin
# delete_paste()
# modify_paste(paste_content)
# Build UI with the new data
demo_paste_content = "200 1 99/99/9999\n1000 2 99/99/9999\n1100 0 99/99/9999\n1200 2 99/99/9999\n1300 1 99/99/9999\n1500 0 99/99/9999\n1600 2 99/99/9999\n1700 1 99/99/9999"
gui(demo_paste_content)








