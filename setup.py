import os
from configparser import ConfigParser
import os
import time
import requests
import os, winshell, win32com.client
from pyunpack import Archive
import wget
isExist = os.path.exists('config.txt')

if isExist == True: 
    Aemail = input("skriv 1/5 av mailadresserna som ska få hembrevet: ")
    Bemail = input("skriv 2/5 av mailadresserna som ska få hembrevet: ")
    Cemail = input("skriv 3/5 av mailadresserna som ska få hembrevet: ")
    Demail = input("skriv 4/5 av mailadresserna som ska få hembrevet: ")
    Eemail = input("skriv 5/5 av mailadresserna som ska få hembrevet: ")
    login = input("skriv din Email som du ska skicka från: ")
    password = input("skriv ditt lösenord för eposten (Detta är nödvändigt): ")


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
        "password": password
    }




#Write the above sections to config.ini file
    with open('config.ini', 'w') as conf:
        config_object.write(conf)
        
isExist = os.path.exists('hembrev.rar')
if isExist == False:
    print("Laddar ner...")
    wget.download("http://hembrev.ga")
# URL of the image to be downloaded is defined as image_url
    print(" ")
    print("Klar!")
print(" ")

print("Installerar...")
user = os.getlogin()
Archive('hembrev.rar').extractall(f"C:/Users/{user}/AppData/Roaming")
# Directory 
desktop = winshell.desktop()
#desktop = r"path to where you wanna put your .lnk file"
path = os.path.join(desktop, 'Hembrev.lnk')
target = rf"C:\Users\{user}\AppData\Roaming\hembrev\hembrev.exe"
icon = "https://iconarchive.com/download/i65094/double-j-design/ravenna-3d/Mail.ico"

shell = win32com.client.Dispatch("WScript.Shell")
shortcut = shell.CreateShortCut(path)
shortcut.IconLocation = icon
shortcut.Targetpath = target
shortcut.save()


# Directory 
directory = "hembrev"
    
# Parent Directory path 
parent_dir = f"C:/Users/{user}/AppData/Roaming/Microsoft/Windows/Start Menu/Programs"
    
# Path 
path = os.path.join(parent_dir, directory) 
    
# Create the directory 
# 'GeeksForGeeks' in 
# '/home / User / Documents' 

try:
    os.mkdir(path) 
except:
    print("")

desktop = f"C:/Users/{user}/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/hembrev"
#desktop = r"path to where you wanna put your .lnk file" 
path = os.path.join(desktop, 'Hembrev.lnk')
target = rf"C:\Users\{user}\AppData\Roaming\hembrev\hembrev.exe"
icon = "https://iconarchive.com/download/i65094/double-j-design/ravenna-3d/Mail.ico"

shell = win32com.client.Dispatch("WScript.Shell")
shortcut = shell.CreateShortCut(path)
shortcut.Targetpath = target
shortcut.IconLocation = icon
shortcut.save()
print(" ")
os.remove("hembrev.rar")
print("klar!")