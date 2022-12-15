import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import configparser

## FILE TO SEND AND ITS PATH
filename = 'hembrev v50.docx'
config = configparser.ConfigParser()
config.read('config.ini')
mail = config.get('login', 'email')
password = config.get('login', 'password')


msg = MIMEMultipart()
msg['From'] = mail
msg['To'] = 'liam@suorsa.se'
msg['Subject'] = 'Report Update'
body = 'Body of the message goes in here'
msg.attach(MIMEText(body, 'plain'))

## ATTACHMENT PART OF THE CODE IS HERE
attachment = open(filename, 'rb')
part = MIMEBase('application', "octet-stream")
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
msg.attach(part)

server = smtplib.SMTP('smtp.office365.com', 587)  ### put your relevant SMTP here
server.ehlo()
server.starttls()
server.ehlo()
server.login(mail, password)  ### if applicable
server.send_message(msg)
server.quit()