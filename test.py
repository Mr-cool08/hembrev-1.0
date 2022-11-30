import os
import configparser
config = configparser.ConfigParser()

isExist = os.path.exists('config.txt')

if isExist  == False:
    Aemail = input("skriv ett av mailadresserna som ska få hembrevet: ")
    Bemail = input("skriv ett av mailadresserna som ska få hembrevet: ")
    Cemail = input("skriv ett av mailadresserna som ska få hembrevet: ")
    Demail = input("skriv ett av mailadresserna som ska få hembrevet: ")
    Eemail = input("skriv ett av mailadresserna som ska få hembrevet: ")
    config.add_section('Emails')
    config.set('Emails', 'Email1', Aemail)
    config.set('Emails', 'Email2', Bemail)
    config.set('Emails', 'Email3', Cemail)
    config.set('Emails', 'Email4', Demail)
    config.set('Emails', 'Email5', Eemail)
# Write the new structure to the new file
    with open(r"config.txt", 'w') as config:
        config.write(config)
isExist = os.path.exists('config.txt')
if isExist == False:
    print("an error acurred")