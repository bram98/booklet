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
from helper_functions import init_folder, copy_file, fuzzy_equal, write_errors, clear_errors 

os.chdir(os.path.dirname(__file__))


#%%
def set_entry(name, column, new_value):
    global data
    if data['Full Name'].str.match( name ).sum() != 1:
        raise Except(f'Could not find name {name}')
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

data_path = Path('responses.xlsx')

set_entry('Kirsten Siebers', 'Photo of yourself', 'KirstenSiebers_Kirsten Siebers.jpg')
set_entry('Amanda van der Sijs', 'Project Description', './manual_files/project_description 86 Amanda van der Sijs.docx;./manual_files/references 86 Amanda van der Sijs.xlsx')

set_entry('Sibylle Schwartmann', 'Project Description', 'project_description 98 Sibylle Schwartmann.docx')
set_entry('Sibylle Schwartmann', 'Project Title', 'Monitoring the Electrooxidation of Lignin Model Compounds with *in situ* ATR-IR spectroscopy')
set_entry('Sibylle Schwartmann', 'Project Type', 'PhD')

set_entry('Susana', 'Project Description', './manual_files/project_description 84 Susana Marin Aguilar.docx')
set_entry('Susana', 'Figure (optional)', 'figure 84 Susana Marin Aguilar.pdf')
set_entry('Susana', 'Photo of yourself', 'portrait 84 Susana Marin Aguilar.tiff')

# append_entry('Amanda van der Sijs', 'Project Description', './manual_files/project_description 86 Amanda van der Sijs.docx;./manual_files/references 86 Amanda van der Sijs.xlsx')
append_file_by_name('Amanda van der Sijs.xlsx', './manual_files', 'response_data/project_descriptions')
append_file_by_name('Amanda van der Sijs.docx', './manual_files', './response_data/project_descriptions')

append_file_by_name('project_description 98 Sibylle Schwartmann.docx', './manual_files', './response_data/project_descriptions')

append_file_by_name('project_description 84 Susana Marin Aguilar.docx', './manual_files', './response_data/project_descriptions')
append_file_by_name('figure 84 Susana Marin Aguilar.pdf', './manual_files', './response_data/figures')
append_file_by_name('portrait 84 Susana Marin Aguilar.tiff', './manual_files', './response_data/portraits')


data_path.replace(data_path.with_stem('responses_old'))
data.to_excel('responses.xlsx')


#%%

'''
Run after rename_files
'''
md = Path()
rd = Path('./response_data')

data = pd.read_excel('responses.xlsx')
data.set_index('ID', inplace=True)




# %%
