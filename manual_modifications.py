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
    
    # Sanity check
    if not name in new_value:
        raise Exception(f'Set_entry_file error: {name} not in {new_value}')
    
    if data['Full Name'].str.match( name ).sum() != 1:
        raise Exception(f'Could not find name {name}')
    person = data[data['Full Name'].str.match( name )].iloc[0]
    
    research_groups = person["Research Group"].split(';')
    if len(research_groups) != 2:
        raise Exception(f'{person.name} {person["Full Name"]} has multiple research groups')
        
    files = new_value.split(';')
    files = [f'../manual_files/{research_groups[0]}/' + file for file in files]
    
    new_value = ';'.join(files)
    
    data.loc[person.name, column] = new_value
    
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
if __name__ == '__main__':
    data = pd.read_excel('responses.xlsx')
    data.set_index('ID', inplace=True)
    
    data_path = Path('responses.xlsx')
    
    columns = data.columns
    print(columns)
#%%
PROJ_DESCRIP = 'Project Description'
CAPTION = u'Figure Caption\xa0(optional)'
TYPE = 'Project Type'
TITLE = 'Project Title'
PORTRAIT = 'Photo of yourself'
FIGURE = 'Figure (optional)'
def manual_modify():
    global data
    '''
    Run before rename_files
    '''
    data = pd.read_excel('responses.xlsx')
    data.set_index('ID', inplace=True)
    
    copy_tree('./manual_files', './response_data/manual_files')
    
    data_path = Path('responses.xlsx')
    
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
    set_entry('George Tierney', 'Project Type', 'Postdoc')
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
    set_entry('Susana', 'Project Type', 'Postdoc')
    set_entry_file('Tjom Arens', 'Project Description', 'project_description 127 Tjom Arens.docx;references 127 Tjom Arens.xlsx')
    set_entry_file('Gerardo Campos-Villalobos', 'Project Description', 'project_description 91 Gerardo Campos-Villalobos.docx;references 91 Gerardo Campos-Villalobos.xlsx')
    set_entry('Gerardo Campos-Villalobos', CAPTION, 'Schematic representation of the process of coarse-graining a system of ligand-stabilized nanoparticles departing from a fine-grained representation.')
    set_entry('Gerardo Campos-Villalobos', TITLE, 'Utilizing Machine Learning for the Bottom-Up Coarse-Graining of Colloidal Systems')
    set_entry('Gerardo Campos-Villalobos', TYPE, 'Postdoc')
    
    '''
    ICC
    '''
    set_entry_file('Bettina Baumgartner', PROJ_DESCRIP, 'project_description 54 Bettina Baumgartner.doc')
    set_entry('Bettina Baumgartner', CAPTION, 'Outline of the experimental setup that comprises a Silicon ATR crystal coated with porous material, which is placed in a flow cell. Gases or liquids can be applied and a window allows top illumination for UV-vis spectroscopic reaction monitoring or initiation of photoreactions. ')
    set_entry('Bettina Baumgartner', TITLE, 'Reaction monitoring in confined spaces using in situ spectroscopy')
    set_entry('Bettina Baumgartner', TYPE, 'Postdoc')
    set_entry_file('Floor Brzesowsky', PROJ_DESCRIP, 'project_description 38 Floor Brzesowsky.docx;references 38 Floor Brzesowsky.xlsx')
    set_entry_file('Albaraa Falodah', PROJ_DESCRIP, 'project_description 60 Albaraa Falodah.docx;references 60 Albaraa Falodah.xlsx')
    set_entry('Albaraa Falodah', CAPTION, 'The evolution of metallocene catalysts for polyolefin polymerization, shown through examples.')
    set_entry_file('Disha Jain', PROJ_DESCRIP, 'project_description 131 Disha Jain.docx;references 131 Disha Jain.xlsx')
    set_entry_file('Joyce Kromwijk', PROJ_DESCRIP, 'project_description 55 Joyce Kromwijk.docx;references 55 Joyce Kromwijk.xlsx')
    set_entry('Joyce Kromwijk', TITLE, 'Two-step Thermochemical CO2 Hydrogenation Toward Aromatics')
    set_entry('Joyce Kromwijk', TYPE, 'PhD')
    set_entry_file('Sebastian Rejman', PROJ_DESCRIP, 'project_description 94 Sebastian Rejman.docx;references 94 Sebastian Rejman.xlsx')
    set_entry_file('Kirsten Siebers', PROJ_DESCRIP, 'project_description 1 Kirsten Siebers.docx;references 1 Kirsten Siebers.xlsx')
    set_entry_file('Kirsten Siebers', PORTRAIT, 'KirstenSiebers_Kirsten Siebers.jpg')
    set_entry_file('Chunning Sun', PROJ_DESCRIP, 'project_description 120 Chunning Sun.docx')
    set_entry('Chunning Sun', TITLE, 'Converting chlorinated products into higher value chemicals')
    set_entry('Chunning Sun', TYPE, 'Postdoc')
    set_entry_file('Xiang Yu', PROJ_DESCRIP, 'project_description 111 Xiang Yu.docx')
    set_entry('Xiang Yu', TITLE, 'Towards an atom-level design and understanding of heterogeneous catalysts for environmental remediations')
    set_entry('Xiang Yu', TYPE, 'PhD')

    
    
    
    data_path.replace(data_path.with_stem('responses_old'))
    data.to_excel('responses.xlsx')
    
if __name__ == '__main__':
    manual_modify()
#%%

'''
Run after rename_files
'''





# %%
