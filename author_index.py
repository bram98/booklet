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
import os

os.chdir(os.path.dirname(__file__))
from helper_functions import group_abbr

#%%
def is_single_letter(string):
    return bool(re.match(r'^[A-Za-z]$', string))

def initial(name):
    return name[0] + '.'

def add_pages(data_sorted) -> None:
    
    # page offsets with respect to last page of previous group
    STARTING_PAGE = 8
    offsets = {
        'CMI': 2,
        'ICC': 2,
        'MCC': 2,
        'NP': 2,
        'OCC': 2,
        'PCC': 2,
    }
    
    # list for bookkeeping
    groups_not_encountered = ['CMI', 'ICC', 'MCC', 'NP', 'OCC', 'PCC']
    
    page = STARTING_PAGE
    for (ID, person) in data_sorted.iterrows():
        group = group_abbr(person['Research Group'])
        
        # print(group)
        if group in groups_not_encountered:
            # new group!
            page += offsets[group]
            groups_not_encountered.remove(group)
        data_sorted.loc[ID, 'Page number'] = str(page)
        # print(data_sorted.loc[ID, 'Page number'])
        # print(person)
        page += 1
        
    
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
# author_index = data[['Full Name', 'E-mail']]
author_index = pd.DataFrame(
    data={
        'Surname, Initials':    list(data['Full Name']),
        'Name':                 list(data['Full Name']),
        'E-mail':               list(data['Email']),
        'Full Name':            list(data['Full Name']),
        },
    # data=[data['Full Name'], data['Name'], data['E-mail'], data['E-mail']],
    # data=np.array(map(lambda c: list(data[c]), ['Full Name', 'Name', 'E-mail', 'E-mail'])).T,
    # columns=['Surname, Initials', 'Name', 'E-mail', 'Page number']
    )

email_initials_regex = re.compile(r'^((?:[a-z]\.)+)')

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
    firstnames_initial_str = ' '.join(firstnames_initials)
    particles_str = ' '.join(particles)
    
    email_initials = email_initials_regex.search(email)
    if not email_initials:
        raise Exception(f'Error parsing initials from {email}')
        
    email_initials = email_initials.groups(1)[0]
    
    manual_initials = False
    if email_initials[0] != firstnames_initial_str.lower()[0] :
        
        print(f'{email_initials:8s} {firstnames_initial_str.lower():5s}', full_name, email)
        match full_name:
            case 'Ivo Vermaire':
                print('iivooo')
                firstnames_initial_str = 'I. A. '
                manual_initials = True
            case 'Kelly Brouwer':
                print('kelz')
                firstnames_initial_str = 'K. J. H. '
                manual_initials = True
            case _:
               pass
   
    if not manual_initials:
        firstnames_initial_str = email_initials.upper().replace('.', '. ')
    
    if firstnames_initial_str.replace(' ','').lower() != email_initials:
        # print(f'{email_initials:8s} {firstnames_initial_str.lower():5s}')
        if email_initials[0] != firstnames_initial_str.lower()[0] :
            print(f'{email_initials:8s} {firstnames_initial_str.lower():5s}', full_name, email)
            # print('!!!', full_name)
    
    author_index.loc[ID, 'Surname, Initials'] = f'{surname}, {firstnames_initial_str} {particles_str}'
    author_index.loc[ID, 'Name'] = firstname

#%%
author_index['Research Group'] = list(data['Research Group'])

# Sort by literal last name, same way as the folders are ordered in responses-conny
data_sorted = author_index.sort_values(
    by='Full Name', 
    kind='stable',
    inplace=False, 
    key=lambda col: col.apply(
            lambda name: name.strip().split()[-1]
        )
    )
data_sorted.sort_values(by='Research Group', kind='stable', inplace=True)
data_sorted['Page number'] = ''

add_pages(data_sorted)

data_sorted.drop(columns='Research Group', inplace=True)


    
#%%
# author_index = 69
print(data_sorted.columns)
writer = pd.ExcelWriter('author_index.xlsx') 
data_sorted.to_excel(writer, sheet_name='sheetName', index=False, na_rep='NaN')

for column in data_sorted:
    column_length = max(data_sorted[column].astype(str).map(len).max(), len(column))
    col_idx = data_sorted.columns.get_loc(column)
    writer.sheets['sheetName'].set_column(col_idx, col_idx, column_length)

writer.close()
# author_index.to_excel('author_index.xlsx')