import os
from tkinter import *


root = Tk()

path = os.path.join("C:\\", "VerificationReports")
isdir = os.path.isdir(path)
if not isdir:
    os.mkdir(path)


path = os.path.join("C:\\VerificationReports", "Analytics(Full)")
isdir = os.path.isdir(path)
if not isdir:
    os.mkdir(path)

path = os.path.join("c:\\VerificationReports", "DownloadDirectory")
isdir = os.path.isdir(path)
if not isdir:
    os.mkdir(path)

path = os.path.join("c:\\VerificationReports", "TaskIngestion")
isdir = os.path.isdir(path)
if not isdir:
    os.mkdir(path)




def firsttimebutton():
    file = open(r"assets\loginInfo.txt", "w+")
    L = [idbox.get() + "\n", passbox.get() + "\n", userbox.get() + "\n", sigbox.get() + "\n", reasonbox.get() + "\n"]
    file.writelines(L)
    file.seek(0)
    file.close()
    root.destroy()

def displaydata():
    file = open(r"assets\loginInfo.txt", "r+")
    a = file.readlines()
    print(a[0].strip())
    print(a[1].strip())
    print(a[2].strip())
    print(a[3].strip())
    print(a[4].strip())

    file.seek(0)
    file.close()


Label(root, text="Please fill these fields following the instructions as accurately as possible.", font=("Verdana", 10)).pack()
Label(root, text="Your www.cozeva.com login ID", font=("Verdana", 10)).pack()
idbox = Entry(root, width=40)
idbox.pack()
Label(root, text="Your www.cozeva.com login Password", font=("Verdana", 10)).pack()
passbox = Entry(root, show="*", width=40)
passbox.pack()
Label(root, text="Enter computer name (example: 'wdey' is my computer name)", font=("Verdana", 10)).pack()
userbox = Entry(root, width=40)
userbox.pack()
Label(root, text="Enter signature (Full name as is in your Cozeva settings)", font=("Verdana", 10)).pack()
sigbox = Entry(root, width=40)
sigbox.pack()
Label(root, text="Enter Reason for login (Designated Production Verification RM)", font=("Verdana", 10)).pack()
reasonbox = Entry(root, width=40)
reasonbox.pack()
Button(root, text="Submit", command=firsttimebutton, font=("Verdana", 10)).pack()
Button(root, text="show existing data", command=displaydata, font=("Verdana", 10)).pack()



root.title("First Time Setup")
root.iconbitmap("assets/icon.ico")
#root.geometry("400x400")
root.mainloop()

