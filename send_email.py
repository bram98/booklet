#%%
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import re
import os
os.chdir(os.path.dirname(__file__))

subject = "[ACTION REQUIRED] Extra info Postgraduate booklet"
# body = "Hello!"
sender = "b.verreussel@uu.nl"
recipients = ["b.verreussel@uu.nl"]
# password = "jvcx pqoj ahtk szah"
password = "yjfhqkmdizzkcbso"

with open('email.txt') as file:
    body = file.read()
    body = re.sub('\n','<br>',body)

def send_email(subject, body, sender, recipients, password):
    msg = MIMEText(body, 'html')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)
    # with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp_server:
    # with smtplib.SMTP_SSL('smtp.office365.coms', 587 ) as smtp_server:
        
    # context=ssl.create_default_context()
    # context = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    with smtplib.SMTP('smtp.office365.com',587, timeout=30) as smtp_server:
        # smtp_server.connect('smtp.office365.com',587)
        smtp_server.ehlo()
        smtp_server.starttls()
        # smtp_server.login(sender, password)
        result = smtp_server.sendmail(sender, recipients, msg.as_string())
    print("Message sent!")
    print(result)
print(body)



send_email(subject, body, sender, recipients, password)