import os
from configparser import ConfigParser
import os
import time
import requests
import os, winshell, win32com.client
from pyunpack import Archive
import wget
from progress.bar import Bar
import sys
import shutil
import urllib.request
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError
req = Request("https://hembrev-1.mrcoolcool.repl.co")
print("testar servrarna...")
try:
    response = urlopen(req)
except HTTPError as e:
    print('Servern funkar inte just nu testa igen senare.')
    print('Error code: ', e.code)
    time.sleep(4)
    sys.exit()
    
except URLError as e:
    print('Servern funkar inte just nu testa igen senare.')
    print('Reason: ', e.reason)
    time.sleep(4)
    sys.exit()
    
else:
    print ('Servrarna funkar :)')
isExist = os.path.exists('config.ini')
user = os.getlogin()
if isExist == False: 
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
    wget.download("https://hembrev-1.mrcoolcool.repl.co")
# URL of the image to be downloaded is defined as image_url
    print(" ")
    print("Klar!")
print(" ")

print("Installerar...")
bar = Bar('Processing', max=26)
for i in range(1):
    Archive('hembrev.rar').extractall(f"C:/Users/{user}/AppData/Roaming")
    bar.next()
    try:
        shutil.move("config.ini", fr"C:/Users/{user}/AppData/Roaming/hembrev")
    except:
        print(" ")
    bar.next()
    desktop = winshell.desktop()
    bar.next()
#desktop = r"path to where you wanna put your .lnk file"
    path = os.path.join(desktop, 'Hembrev.lnk')
    bar.next()
    target = rf"C:\Users\{user}\AppData\Roaming\hembrev\hembrev.exe"
    bar.next()
    icon = f"C:/Users/{user}/AppData/Roaming/hembrev/Double-J-Design-Ravenna-3d-Mail.ico"
    
    bar.next()
    shell = win32com.client.Dispatch("WScript.Shell")
    bar.next()
    shortcut = shell.CreateShortCut(path)
    bar.next()
    shortcut.IconLocation = icon
    bar.next()
    shortcut.Targetpath = target
    bar.next()
    shortcut.save()
    

    bar.next()
    directory = "hembrev"
    bar.next()
    
# Parent Directory path 
    parent_dir = f"C:/Users/{user}/AppData/Roaming/Microsoft/Windows/Start Menu/Programs"
    bar.next()
    
# Path 
    path = os.path.join(parent_dir, directory) 
    bar.next()
    try:
        os.mkdir(path) 
    except:
        print("")
    bar.next()
    desktop = f"C:/Users/{user}/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/hembrev"
    bar.next()
    #desktop = r"path to where you wanna put your .lnk file" 
    path = os.path.join(desktop, 'Hembrev.lnk')
    bar.next()
    target = rf"C:\Users\{user}\AppData\Roaming\hembrev\hembrev.exe"
    bar.next()
    icon = rf"C:\Users\{user}\AppData\Roaming\hembrev\Double-J-Design-Ravenna-3d-Mail.ico"
    
    bar.next()
    shell = win32com.client.Dispatch("WScript.Shell")
    bar.next()
    shortcut = shell.CreateShortCut(path)
    bar.next()
    shortcut.Targetpath = target
    bar.next()
    shortcut.IconLocation = icon
    bar.next()
    shortcut.save()
    bar.next()
    print(" ")
    bar.next()
    os.remove("hembrev.rar")
    bar.next()
bar.finish()

# Directory 





# Directory 

    
# Create the directory 
# 'GeeksForGeeks' in 
# '/home / User / Documents' 







print("klar!")
