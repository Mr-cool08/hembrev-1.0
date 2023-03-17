import tkinter as tk
from tkinter import messagebox
from configparser import ConfigParser
from cryptography.fernet import Fernet


import time
import requests
import os

from tqdm import tqdm
from windtalker import SymmetricCipher
import os
import zipfile
import sys
import shutil
import urllib.request
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError

def encrypt():
    c = SymmetricCipher(password="password")
    try:
        c.encrypt_file("password.txt")
        os.remove("password.txt")
    except OSError:
        os.remove("password-encrypted.txt")
        
def submit():
    global Aemail
    global Bemail
    global Cemail
    global Demail
    global Eemail
    global login
    global password
    Aemail = email1_var.get()
    Bemail = email2_var.get()
    Cemail = email3_var.get()
    Demail = email4_var.get()
    Eemail = email5_var.get()
    login = login_var.get()
    password = password_var.get()
    config_object = ConfigParser()
    config_object["Emails"] = {
        "email1": Aemail,
        "email2": Bemail,
        "email3": Cemail,
        "email4": Demail,
        "email5": Eemail
    }
    config_object["login"] = {
        "email": login,
    }
    with open('config.ini', 'w') as conf:
        config_object.write(conf)
    with open('password.txt', 'w') as f:
        f.write(password)
    encrypt()
    os.startfile('hembrev.exe')
    messagebox.showinfo("Success", "Program started!")

root = tk.Tk()
root.title("Program Input")
root.geometry("400x300")

email1_var = tk.StringVar()
email2_var = tk.StringVar()
email3_var = tk.StringVar()
email4_var = tk.StringVar()
email5_var = tk.StringVar()
login_var = tk.StringVar()
password_var = tk.StringVar()

entry = tk.Entry(root, textvariable=email1_var)
email1_label = tk.Label(root, text="Email 1: ")
email1_entry = tk.Entry(root, textvariable=email1_var)
email2_label = tk.Label(root, text="Email 2: ")
email2_entry = tk.Entry(root, textvariable=email2_var)
email3_label = tk.Label(root, text="Email 3: ")
email3_entry = tk.Entry(root, textvariable=email3_var)
email4_label = tk.Label(root, text="Email 4: ")
email4_entry = tk.Entry(root, textvariable=email4_var)
email5_label = tk.Label(root, text="Email 5: ")
email5_entry = tk.Entry(root, textvariable=email5_var)
login_label = tk.Label(root, text="School Email: ")
login_entry = tk.Entry(root, textvariable=login_var)
password_label = tk.Label(root, text="Password: ")
password_entry = tk.Entry(root, textvariable=password_var, show="*")
submit_button = tk.Button(root, text="Submit", command=submit)

email1_label.grid(row=0, column=0, padx=5, pady=5)
email1_entry.grid(row=0, column=1, padx=5, pady=5)
email2_label.grid(row=1, column=0, padx=5, pady=5)
email2_entry.grid(row=1, column=1, padx=5, pady=5)
email3_label.grid(row=2, column=0, padx=5, pady=5)
email3_entry.grid(row=2, column=1, padx=5, pady=5)
email4_label.grid(row=3, column=0, padx=5, pady=5)
email4_entry.grid(row=3, column=1, padx=5, pady=5)
email5_label.grid(row=4, column=0, padx=5, pady=5)
email5_entry.grid(row=4, column=1, padx=5, pady=5)
login_label.grid(row=5, column=0, padx=5, pady=5)
login_entry.grid(row=5, column=1, padx=5, pady=5)
password_label.grid(row=6, column=0, padx=5, pady=5)
password_entry.grid(row=6, column=1, padx=5, pady=5)
submit_button.grid(row=7, column=1, padx=5, pady=5)

root.mainloop()
