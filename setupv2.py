from configparser import ConfigParser
from cryptography.fernet import Fernet
import os
from progress.bar import Bar
import time
import requests
import os, winshell, win32com.client
from pyunpack import Archive
import wget
from tqdm import tqdm
from windtalker import SymmetricCipher
import os
import zipfile
import sys
import shutil
import urllib.request
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError
isExist = os.path.exists('config.ini')
user = os.getlogin()
if isExist == False:
    print("Om du inte har flera epostadresser att skicka till lämna blankt")
    Aemail = input("skriv 1/5 av mailadresserna som ska få hembrevet: ")
    Bemail = input("skriv 2/5 av mailadresserna som ska få hembrevet: ")
    Cemail = input("skriv 3/5 av mailadresserna som ska få hembrevet: ")
    Demail = input("skriv 4/5 av mailadresserna som ska få hembrevet: ")
    Eemail = input("skriv 5/5 av mailadresserna som ska få hembrevet: ")
    login = input("skriv din Email som du ska skicka från: ")


    config_object = ConfigParser()
    
#Assume we need 2 sections in the config file, let's call them USERINFO and SERVERCONFIG
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




#Write the above sections to config.ini file
    with open('config.ini', 'w') as conf:
        config_object.write(conf)
password = input("skriv ditt lösenord för eposten (Detta är nödvändigt): ")        




with open('password.txt', 'w') as f:
        f.write(password)
        
def encrypt():
    c = SymmetricCipher(password="password")
    try:
        c.encrypt_file("password.txt")
        os.remove("password.txt")
    except OSError:
        os.remove("password-encrypted.txt")
        
        
encrypt()     
        
        
        
        
        
        
        
