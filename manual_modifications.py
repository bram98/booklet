#%%
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import PIL
from  glob import glob
import os
import shutil
from unidecode import unidecode
from urllib.parse import urlparse, parse_qs, unquote
import re
from tqdm import tqdm
import json
from fuzzywuzzy import fuzz
from pathlib import Path
from distutils.dir_util import copy_tree

os.chdir(os.path.dirname(__file__))
from helper_functions import init_folder, copy_file, fuzzy_equal, write_errors, clear_errors 



#%%
def set_entry(name, column, new_value):
    global data
    
    if data['Full Name'].str.match( name ).sum() != 1:
        raise Exception(f'Could not find name {name}')
    data.loc[ data['Full Name'].str.match( name ), column ] = new_value
    
def set_entry_file(name, column, new_value):
    global data
    files = new_value.split(';')
    files = ['../manual_files/' + file for file in files]
    new_value = ';'.join(files)
    # print(new_value)
    if data['Full Name'].str.match( name ).sum() != 1:
        raise Exception(f'Could not find name {name}')
    data.loc[ data['Full Name'].str.match( name ), column ] = new_value
    
def append_entry(name, column, new_value):
    global data
    data.loc[ data['Full Name'] == name, column ] += new_value
    
def replace_file_by_name(name, src_dir, target_dir):
    src_files = glob(src_dir + '/*')
    target_files = glob(target_dir + '/*')
    src_files = [file for file in src_files if name in file]
    target_files = [file for file in target_files if name in file]
    if len(src_files) != 1:
        raise Exception(f'src_files: {src_files}')
    if len(target_files) != 1:
        raise Exception(f'target_files: {target_files}')
    src_file = src_files[0]
    target_file = target_files[0]
    
    shutil.copy(src_file, target_file)
    print(f'Moved {src_file} to {target_file}')
    
def append_file_by_name(name, src_dir, target_file):
    src_files = glob(src_dir + '/*')
    # target_files = glob(target_dir + '/*')
    src_files = [file for file in src_files if name in file]
    # target_files = [file for file in target_files if name in file]
    if len(src_files) != 1:
        raise Exception(f'src_files: {src_files}')
    # if len(target_files) != 1:
    #     raise Exception(f'target_files: {src_files}')
    src_file = src_files[0]
    # target_file = target_files[0]
    
    shutil.copy(src_file, target_file)
    print(f'Moved {src_file} to {target_file}')
#%%
data = pd.read_excel('responses.xlsx')
data.set_index('ID', inplace=True)

data_path = Path('responses.xlsx')

columns = data.columns
print(columns)
#%%

'''
Run before rename_files
'''
data = pd.read_excel('responses.xlsx')
data.set_index('ID', inplace=True)

copy_tree('./manual_files', './response_data/manual_files')

data_path = Path('responses.xlsx')

set_entry_file('Kirsten Siebers', 'Photo of yourself', 'KirstenSiebers_Kirsten Siebers.jpg')
set_entry_file('Amanda van der Sijs', 'Project Description', 'project_description 86 Amanda van der Sijs.docx;references 86 Amanda van der Sijs.xlsx')

set_entry_file('Sibylle Schwartmann', 'Project Description', 'project_description 98 Sibylle Schwartmann.docx')
set_entry('Sibylle Schwartmann', 'Project Title', 'Monitoring the Electrooxidation of Lignin Model Compounds with *in situ* ATR-IR spectroscopy')
set_entry('Sibylle Schwartmann', 'Project Type', 'PhD')


'''MCC'''

set_entry_file('Qijun Che', 'Project Description', 'project_description 78 Qijun Che.doc;references 78 Qijun Che.xlsx')
set_entry_file('Maaike Vink', 'Project Description', 'project_description 34 Maaike Vink-van Ittersum.doc;references 34 Maaike Vink-van Ittersum.xlsx')
set_entry('Maaike Vink', 'Project Title', 'New electrodes for CO2 Conversion: Nanomorphology and Composition')
set_entry('Maaike Vink', 'Figure Caption (optional)', 'a) Porous Ag prepared via AgAl dealloying; b) Porous Ag prepared via templating with PMMA spheres')
set_entry('Maaike Vink', 'Project Type', 'PhD')
set_entry_file('Claudia Keijzer', 'Project Description', 'project_description 6 Claudia Keijzer.doc;references 6 Claudia Keijzer.xlsx')
set_entry('Claudia Keijzer', 'Project Type', 'PhD')
set_entry('Claudia Keijzer', 'Project Title', 'Ultra-pure and Structured Supports for Silver Catalysts for Ethylene Epoxidation')
set_entry_file('Komal N. Patil', 'Figure (optional)', 'figure 14 Komal N. Patil.jpg')
set_entry_file('Marta', 'Project Description', 'project_description 102 Marta Perxes Perich.docx;references 102 Marta Perxes Perich.xlsx')
set_entry_file('Yuang Piao', 'Project Description', 'project_description 68 Yuang Piao.doc;references 68 Yuang Piao.xlsx')
set_entry('Yuang Piao', 'Project Title', 'Bifunctional catalysts convert syngas to dimethyl ether')
set_entry('Yuang Piao', 'Project Type', 'PhD')
set_entry_file('Henrik Rodenburg', 'Project Description', 'project_description 5 Henrik Rodenburg.docx;references 5 Henrik Rodenburg.xlsx')
set_entry_file('Suzan Schoemaker', 'Project Description', 'project_description 21 Suzan Schoemaker.docx;references 21 Suzan Schoemaker.xlsx')
set_entry('Suzan Schoemaker', 'Figure Caption (optional)', 'Main steps in methane decomposition. 1) Carbon supply: Methane adsorption and dissociation. Hz gas is released. 2) Carbon transport: C-atoms diffuse through or over the catalyst. 3) Formation: carbon nanostructures are formed. 4) Deactivation: formation of defective carbon lavers at the surface.')
set_entry_file('George Tierney', 'Project Description', 'project_description 47 George Tierney.docx;references 47 George Tierney.xlsx')
set_entry('George Tierney', 'Project Type', 'postdoc')
set_entry_file('George Tierney', 'Figure (optional)', 'figure 47 George Tierney.jpg')
set_entry_file('Zixiong Wei', 'Project Description', 'project_description 123 Zixiong Wei.docx;references 123 Zixiong Wei.xlsx')
set_entry('Zixiong Wei', 'Project Type', 'PhD')
set_entry('Zixiong Wei', 'Project Title', 'Machine Learning-based Multiscale Mechanical Simulations of Li/LLZO Interface in Solid-State Batteries for Improved Performance and Reliability')
set_entry_file('Karan Kotalgi', 'Figure (optional)', 'figure 130 Karan Kotalgi.jpg')

'''
SCMB
'''
set_entry_file('Susana', 'Project Description', 'project_description 84 Susana Marin Aguilar.docx')
set_entry_file('Susana', 'Figure (optional)', 'figure 84 Susana Marin Aguilar.pdf')
set_entry_file('Susana', 'Photo of yourself', 'portrait 84 Susana Marin Aguilar.tiff')
set_entry('Susana', 'Project Title', 'Simulations of colloidal particles with anisotropic interactions and shapes')
set_entry('Susana', 'Project Type', 'postdoc')
set_entry_file('Tjom Arens', 'Project Description', 'project_description 127 Tjom Arens.docx;references 127 Tjom Arens.xlsx')
# # append_entry('Amanda van der Sijs', 'Project Description', './manual_files/project_description 86 Amanda van der Sijs.docx;./manual_files/references 86 Amanda van der Sijs.xlsx')
# append_file_by_name('Amanda van der Sijs.xlsx', './manual_files', 'response_data/project_descriptions')
# append_file_by_name('Amanda van der Sijs.docx', './manual_files', './response_data/project_descriptions')

# append_file_by_name('project_description 98 Sibylle Schwartmann.docx', './manual_files', './response_data/project_descriptions')

# append_file_by_name('project_description 84 Susana Marin Aguilar.docx', './manual_files', './response_data/project_descriptions')
# append_file_by_name('figure 84 Susana Marin Aguilar.pdf', './manual_files', './response_data/figures')
# append_file_by_name('portrait 84 Susana Marin Aguilar.tiff', './manual_files', './response_data/portraits')


data_path.replace(data_path.with_stem('responses_old'))
data.to_excel('responses.xlsx')


#%%

'''
Run after rename_files
'''
# md = Path()
# rd = Path('./response_data')

# data = pd.read_excel('responses.xlsx')
# data.set_index('ID', inplace=True)




# %%
