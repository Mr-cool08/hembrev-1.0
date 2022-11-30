import os
from configparser import ConfigParser


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