import os
from configparser import ConfigParser
import os
import time
import requests
import os, winshell, win32com.client
from pyunpack import Archive
import wget
from progress.bar import Bar


import urllib.request
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError
req = Request("https://hembrev.ga")
print("testar servrarna...")
try:
    response = urlopen(req)
except HTTPError as e:
    print('Servern funkar inte just nu testa igen senare.')
    print('Error code: ', e.code)
    
except URLError as e:
    print('Servern funkar inte just nu testa igen senare.')
    print('Reason: ', e.reason)
    
else:
    print ('Servrarna funkar :)')
isExist = os.path.exists('config.txt')
user = os.getlogin()
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
    wget.download("https://hembrev.ga")
# URL of the image to be downloaded is defined as image_url
    print(" ")
    print("Klar!")
print(" ")

print("Installerar...")
bar = Bar('Processing', max=20)
for i in range(20):
    Archive('hembrev.rar').extractall(f"C:/Users/{user}/AppData/Roaming")
    bar.next()
    os.replace("config.ini", f"C:/Users/{user}/AppData/Roaming/hembrev/config.ini")
    bar.next()
    desktop = winshell.desktop()
#desktop = r"path to where you wanna put your .lnk file"
    path = os.path.join(desktop, 'Hembrev.lnk')
    target = rf"C:\Users\{user}\AppData\Roaming\hembrev\hembrev.exe"
    icon = f"C:/Users/{user}/AppData/Roaming/hembrev/Double-J-Design-Ravenna-3d-Mail.ico"
    bar.next()
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(path)
    shortcut.IconLocation = icon
    shortcut.Targetpath = target
    shortcut.save()

    bar.next()
    
    bar.next()
    
    bar.next()
    
    bar.next()
    
    bar.next()
bar.finish()

# Directory 





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
icon = rf"C:\Users\{user}\AppData\Roaming\hembrev\Double-J-Design-Ravenna-3d-Mail.ico"

shell = win32com.client.Dispatch("WScript.Shell")
shortcut = shell.CreateShortCut(path)
shortcut.Targetpath = target
shortcut.IconLocation = icon
shortcut.save()
print(" ")
os.remove("hembrev.rar")
print("klar!")
