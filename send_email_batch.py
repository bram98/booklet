#%%
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import re
import json
import os
os.chdir(os.path.dirname(__file__))

import pandas as pd
#%%
subject = "[ACTION REQUIRED] Personal improvements for submission Postgraduate Booklet"
# body = "Hello!"
sender = "b.verreussel@gmail.com"
recipients = ["b.verreussel@gmail.com"]
password = "jvcx pqoj ahtk szah"
# password = "yjfhqkmdizzkcbso"

def send_email(subject, body, sender, recipients, password):
    msg = MIMEText(body, 'html')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp_server:

        smtp_server.login(sender, password)
        smtp_server.sendmail(sender, recipients, msg.as_string())
    print("Message sent!")
    # print(result)

#%%
# Read in the data as a pandas dataframe
data = pd.read_excel('responses.xlsx')
data.set_index('ID', inplace=True)

with open('error_data.json') as file:
    error_dict = json.load(file)

#%%
i=0

for (ID, person) in error_dict.items():
    msg = (f'Dear {person["name"]},\n\n'
           'The following issues have been detected with your form submission for the Postgraduate Booklet:\n\n'
           )
    
    # Assume innocence until proven otherwise
    good_person = True
    
    for err_type in ['portrait_errors', 'figure_errors', 'proj_description_errors', 
                     'references_errors', 'name_errors']:
        
        if len(person[err_type]) > 0:
             msg += '\n\n'.join(person[err_type])
             msg += '\n\n'
             good_person = False
    
    msg += ''
    if good_person:
        # print(person['name'], ID, 'is a good person!')
        msg += ('No issues have been detected with your response. Thank you for your '
                'cooperation :)\n\n')

    print()
    i+=1
    if i >= 20:
        break
    msg += ('If you have questions or if you do not agree with this error message,'
            ' you can send an e-mail to b.verreussel@uu.nl.\n\n'
            'Best regards,\nOn behalf of the DAC,\n\nBram Verreussel')
    print(msg)
    msg = re.sub('\n','<br>', msg )
    recipients = [person['email']]
    # if good_person:
    send_email(subject, msg, sender, recipients, password)
        # break

