# -*- coding: utf-8 -*-
"""
Created on Thu Feb  1 10:26:00 2024

@author: 5914787
"""
#%%
import numpy as np
import pandas as pd
from unidecode import unidecode
import re

#%%
def is_single_letter(string):
    return bool(re.match(r'^[A-Za-z]$', string))

def initial(name):
    return name[0] + '.'
#%%
# data.to_excel('responses.xlsx')

data = pd.read_excel('responses.xlsx')
data.set_index('ID', inplace=True)

#%%
# author_index = pd.DataFrame(
#     data=[ data['Full Name'], data['Name'], data['Email'] ],
#     columns=['Surname, Initials','Name', 'E-mail', 'Page number']
#     )
# author_index = data.filter(['Full Name', 'Name', 'E-mail'])
# author_index = author_index.rename(
#     columns={
#         'Full Name': 'Surname, Initials',
#         })
# author_index = pd.DataFrame(
#     # data=[data['Full Name'], data['Name'], data['E-mail'], data['E-mail']],
#     data=np.array(map(lambda c: list(data[c]), ['Full Name', 'Name', 'E-mail', 'E-mail'])).T,
#     columns=['Surname, Initials', 'Name', 'E-mail', 'Page number']
#     )

author_index = pd.DataFrame(
    data={
        'Surname, Initials':    list(data['Full Name']),
        'Name':                 list(data['Full Name']),
        'E-mail':               list(data['Email']),
        'Page number':          ['']*len(data),
        },
    # data=[data['Full Name'], data['Name'], data['E-mail'], data['E-mail']],
    # data=np.array(map(lambda c: list(data[c]), ['Full Name', 'Name', 'E-mail', 'E-mail'])).T,
    # columns=['Surname, Initials', 'Name', 'E-mail', 'Page number']
    )

for (ID, person) in author_index.iterrows():
    full_name = person['Surname, Initials']
    full_name = re.sub(r'dr\.\ ?','',full_name, flags=re.I)     # remove Dr. from name
    full_name = re.sub(r'\.', '', full_name)                    # Remove periods from name
    email = person['E-mail']
    
    name_parts = re.split(r'\ ' ,full_name)                   # split on spaces or dashes
    # name_parts = [name_part for name_part in name_parts if not name_part is None]
    
    # tries to determine surnames by looking at presence in email adress. Loops from right to left
    surnames = []
    simple_email = email.replace('-','')
    for name_part in name_parts[::-1]:   
        simple_name_part = name_part.replace('-','')                       
        simple_name_part = unidecode(simple_name_part).lower()
        if simple_name_part in simple_email and not is_single_letter(simple_name_part):
            surnames.insert(0, name_part)
            simple_email = re.sub(simple_name_part, '', simple_email)
        else:
            break
    
    # Sort out particles like van, de etc.
    surnames_wout_particles = surnames.copy()
    particles = []
    for surname in surnames:
        if surname in ['van', 'de', 'der', 'ter', 'Del', 'den']:
            surnames_wout_particles = surnames_wout_particles[1:]
            particles.append(surname)
        else:
            break
    firstnames = [name_part for name_part in name_parts if not name_part in surnames]
    surnames = surnames_wout_particles.copy()
    
    # This one guy did not type his surname in full name
    if len(surnames) == 0:
        surnames = [re.search(r'\w{3,}', email).group().capitalize()]
    
    # Pretty print everything
    surname = ' '.join(surnames)
    # firstname = ' '.join(firstnames)
    firstname = firstnames[0]
    firstnames_initials = [initial(firstname) for firstname in firstnames]
    firstname_initials = ' '.join(firstnames_initials)
    particles_str = ' '.join(particles)
    
    author_index.loc[ID, 'Surname, Initials'] = f'{surname}, {firstname_initials} {particles_str}'
    author_index.loc[ID, 'Name'] = firstname
    
    
    
writer = pd.ExcelWriter('author_index.xlsx') 
author_index.to_excel(writer, sheet_name='sheetName', index=False, na_rep='NaN')

for column in author_index:
    column_length = max(author_index[column].astype(str).map(len).max(), len(column))
    col_idx = author_index.columns.get_loc(column)
    writer.sheets['sheetName'].set_column(col_idx, col_idx, column_length)

writer.close()
# author_index.to_excel('author_index.xlsx')