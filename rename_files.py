#%%
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import PIL
import glob
import os
import shutil
from unidecode import unidecode
from urllib.parse import urlparse, parse_qs, unquote
import re
from tqdm import tqdm

#%%
def init_folder(folder):
    '''
    Creates folder if it does not exist and if it exists, it will delete all files inside.
    '''
    os.makedirs(f'./response_data/{folder}',exist_ok=True)
    print(f'Created folder {folder}')
    files = glob.glob('./response_data/{folder}/*')
    for file in files:
        os.remove(file)
def inspect_names():
    '''
    see how the names compare to the ones provided by the university
    '''
    for(i, person) in data.iterrows():
        name = person['Full Name']
        uuname = person['Name']
        if name != uuname:
            print(f'Name from form:\t{name}\nName from uu:\t{uuname}')

def extract_file_solis_url(url):
    '''
    Extract the file name from the solislink as provided in the response form
    '''
    u = urlparse(url)
    if len(u.query) > 0 :
        # print(len(u.query))
        q = parse_qs( u.query ) 
        try:
            return q['file'][0]
        except:
            return q['name'][0].replace("'","")
    else:
        return os.path.basename(u.path)
    
def copy_file(src_folder, dest_folder, old_filename, new_filename, name):
    '''
    Copies a file from src_folder/old_filename to dest_folder/new_filename {name}.oldextension
    '''
    try:
        filename, file_extension = os.path.splitext(old_filename)
        # if file_extension == '.ai':
        #     file_extension = '.jpg' # I manually converted .ai to .jpg
        src_file_name = f'./response_data/{src_folder}/{filename}{file_extension}'
        dest_file_name = f'./response_data/{dest_folder}/{new_filename} {name}{file_extension}'
        
        shutil.copy(src_file_name, dest_file_name)
    except Exception as e:
        print(f'Error! Name:{name}\t File name:{old_filename}')
        raise e
#%%
# Read in the data as a pandas dataframe
data = pd.read_excel('responses.xlsx')

#%%
# Create folders. Warning: files inside will be deleted
    
init_folder('portraits_renamed')
init_folder('project_descriptions_renamed')
init_folder('references_renamed')
init_folder('figures_renamed')

print('Copying portraits, project descriptions, references and figures')
for(i, person) in tqdm(data.iterrows(), total=len(data)):
    name = str(i) + ' ' + person['Full Name']
    
    '''
    Portraits
    '''
    portrait_file = os.path.split(person['Photo of yourself'])[1]
    portrait_file = unquote(portrait_file)
    copy_file('portraits','portraits_renamed', portrait_file, 'portrait', name)
    
    '''
    Project Description and References
    '''
    proj_description_files = person['Project Description']
    proj_description_files = proj_description_files.split(';')
    proj_description_files = [extract_file_solis_url(url) for url in proj_description_files]

    if not 1 <= len(proj_description_files) <= 2:
        raise Exception(f'Found {len(proj_description_files)} files in project description. Expected 1 or 2.')
    reference_regex = re.compile('ref', flags=re.I)
    if len(proj_description_files)== 1:
        # no references
        proj_description_file = proj_description_files[0]
    elif len(proj_description_files) == 2:
        # references
        if not reference_regex.search(proj_description_files[0]) is None:
            proj_description_file = proj_description_files[1]
            ref_file = proj_description_files[0]
        elif not reference_regex.search(proj_description_files[1]) is None:
            proj_description_file = proj_description_files[0]
            ref_file = proj_description_files[1]
        else:
            print('No references found! ', proj_description_files)
        ref_file = unquote(ref_file)
        
    proj_description_file = unquote(proj_description_file)  
    copy_file('project_descriptions',
              'project_descriptions_renamed',
              proj_description_file, 
              'project_description',
              name)
    copy_file('project_descriptions',
              'references_renamed',
              ref_file, 
              'references',
              name)
    '''
    Figures
    '''
    fig_files = person['Figure (optional)']
    if fig_files is not np.nan:
        # fig_file = os.path.split(fig_files)[1]
        fig_file = extract_file_solis_url(fig_files)
        fig_file = unquote(fig_file)
        copy_file('figures',
              'figures_renamed',
              fig_file, 
              'figure',
              name)

print('Finished copying portraits, project descriptions, references and figures')


