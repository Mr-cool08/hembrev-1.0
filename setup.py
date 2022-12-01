import os
from configparser import ConfigParser
import os

import requests

isExist = os.path.exists('config.txt')


Aemail = input("skriv ett av mailadresserna som ska få hembrevet: ")
Bemail = input("skriv ett av mailadresserna som ska få hembrevet: ")
Cemail = input("skriv ett av mailadresserna som ska få hembrevet: ")
Demail = input("skriv ett av mailadresserna som ska få hembrevet: ")
Eemail = input("skriv ett av mailadresserna som ska få hembrevet: ")



config_object = ConfigParser()

#Assume we need 2 sections in the config file, let's call them USERINFO and SERVERCONFIG
config_object["Emails"] = {
    "email1": Aemail,
    "email2": Bemail,
    "email3": Cemail,
    "email4": Demail,
    "email5": Eemail
}



#Write the above sections to config.ini file
with open('config.ini', 'w') as conf:
    config_object.write(conf)
    
    
    
    
    
    
print("downloading...")
image_url = "https://download1084.mediafire.com/sytyvyxs99fg/vevoqm8z5tip8vq/hembrev.exe"
  
# URL of the image to be downloaded is defined as image_url
r = requests.get(image_url) # create HTTP response object
  
# send a HTTP request to the server and save
# the HTTP response in a response object called r
with open("hembrev.exe",'wb') as f:
  
    # Saving received content as a png file in
    # binary format
  
    # write the contents of the response (r.content)
    # to a new file in binary mode.
    f.write(r.content)
print("done")

os.system("hembrev.exe")
