#%%
import numpy as np
import pandas as pd
import PIL
from  glob import glob
import os
import shutil
from unidecode import unidecode
from urllib.parse import urlparse, parse_qs, unquote
import re
from tqdm import tqdm
import json
from pandas.api.types import is_integer_dtype

os.chdir(os.path.dirname(__file__))

#%%
def init_folder(folder):
    '''
    Creates folder if it does not exist and if it exists, it will delete all files inside.
    '''
    os.makedirs(f'./response_data/{folder}',exist_ok=True)
    print(f'Created folder {folder}')
    files = glob(f'./response_data/{folder}/*')
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
def separate_projdescription_and_references(proj_description_files, ref_regex):
    ref_file = ''
    
    if not reference_regex.search(proj_description_files[0]) is None:
        proj_description_file = proj_description_files[1]
        ref_file = proj_description_files[0]
        has_references = True
    elif not reference_regex.search(proj_description_files[1]) is None:
        proj_description_file = proj_description_files[0]
        ref_file = proj_description_files[1]
        has_references = True
    else:
        print('No references found! ', proj_description_files)
        has_references = False
    
    ref_file = unquote(ref_file)
    return proj_description_file, ref_file, has_references

def create_errors_file():
    print('Creating error file...')
    error_dict = {}
    for(i, person) in data.iterrows():
        error_dict[i] = {
                'name': unidecode(str(person['Full Name'])),
                'email': person['E-mail'],
                'portrait_errors': [],
                'proj_description_errors': [],
                'references_errors': [],
                'name_errors': []
            }
    with open('error_data.json','w') as file:
        json.dump(error_dict, file, indent=4)
        
def write_errors(err_type, id_list, err_list):
    """Writes errors in data to error_data.json. This helps in creating an e-mail request for improving responses.

    Parameters
    ----------
    err_type: str
        error type. Must be one of 'portrait', 'proj_description', 'references' or 'name'
    id_list: list(int)
        list of all people with errors specified by their ID
    err_list: list(str)
        list of strings that will appear in literally in the final email. Same order as id_list
    
    Returns
    -------
        None
    """
    if len(id_list)==0:
        return
    
    with open('error_data.json') as file:
        error_dict = json.load(file)
        for (i, person_id) in enumerate(id_list):
            
            # list of errors for this person currently in error_dict
            person_err_list = error_dict[str(person_id)][f'{err_type}_errors']
            
            # checks if error is already present (prevents writing same error twice)
            if not any( err_list[i] in err for err in person_err_list ):
                person_err_list.append(err_list[i])
    
    with open('error_data.json', 'w') as file:
        json.dump(error_dict, file, indent=4)    

def write_errors_in_name():
    id_list = []
    err_list = []
    for (i, person) in data.iterrows():
        if 'Mengwei' in person['Name']:
            # id_list.append( int(person['ID']) )
            err_list.append(
                (f'There appears to be a typo in your name. Name from database: {person["Name"]},'
                       f' Name provided by the form: {person["Full Name"]}. The name from the form will be used. '
                       'Please modify your response with the correct name')
                )
    
            
    write_errors('name', id_list, err_list)    

    print('Found 1 name error(s)')

def write_errors_references(person_id, excel_file):
    id_list = []
    err_list = []
    df = pd.read_excel(excel_file)

    if df.columns.size != 6:
        id_list.append(person_id)
        err_list.append('Please do not modify the number of columns in the references excel sheet.')
        return  # Cannot process any further if the number of columns isn't even correct
    if not is_integer_dtype(df.iloc[:,0]):
        id_list.append(person_id)
        err_list.append('The first column with reference numbers does not contain (only) numbers.')
    if not is_integer_dtype(df.iloc[:,3]):
        id_list.append(person_id)
        err_list.append('The fourth column with journal volume/issue does not contain (only) numbers.')
    if not is_integer_dtype(df.iloc[:,5]):
        id_list.append(person_id)
        err_list.append('The sixth column with publication year does not contain (only) numbers.')
    for i in range(df.columns.size):
        if np.any(df.iloc[:,i].isna()):
            id_list.append(person_id)
            err_list.append(f'Column {i+1} contains some empty cells')

    write_errors('references', id_list, err_list)

#%%
# Read in the data as a pandas dataframe
data = pd.read_excel('responses.xlsx')
data.set_index('ID', inplace=True)
#%%
MODIFY_FILES = True    # If false, will not modify files. Use for debugging

# Create folders. Warning: files inside will be deleted
if MODIFY_FILES:
    create_errors_file()
    init_folder('portraits_renamed')
    init_folder('project_descriptions_renamed')
    init_folder('references_renamed')
    init_folder('figures_renamed')
    

print('Copying portraits, project descriptions, references and figures')
# for(i, person) in tqdm(data.iterrows(), total=1):
for(ID, person) in tqdm(data.iterrows(), total=len(data)):
    name = f"{ID} {person['Full Name']}" 
    name = unidecode(name)
    
    '''
    Portraits
    '''
    portrait_file = os.path.split(person['Photo of yourself'])[1]
    portrait_file = unquote(portrait_file)
    if MODIFY_FILES:
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
    
    # sadly, one person did not upload a project description and I have to check for this
    has_proj_description = False        
    has_references = False
    
    if len(proj_description_files)== 1:
        if reference_regex.search(proj_description_files[0]) is None:
            # no references
            proj_description_file = proj_description_files[0]
            has_proj_description = True
            has_references = False
        else: 
            # no project description (sigh)
            ref_file = proj_description_files[0]
            has_proj_description = False
            has_references = True
    elif len(proj_description_files) == 2:
        
        # references
        proj_description_file, ref_file, has_references = separate_projdescription_and_references(
            proj_description_files, 
            reference_regex)
        has_proj_description = True

    proj_description_file = unquote(proj_description_file)  
    
    if MODIFY_FILES:
        if has_proj_description:
            copy_file('project_descriptions',
                      'project_descriptions_renamed',
                      proj_description_file, 
                      'project_description',
                      name)
        if has_references:
            copy_file('project_descriptions',
                      'references_renamed',
                      ref_file, 
                      'references',
                      name)
            write_errors_references(
                ID,
                f'./response_data/project_descriptions/{ref_file}'
            )
    '''
    Figures
    '''
    fig_files = person['Figure (optional)']
    if fig_files is not np.nan:
        # fig_file = os.path.split(fig_files)[1]
        fig_file = extract_file_solis_url(fig_files)
        fig_file = unquote(fig_file)
        if MODIFY_FILES:
            copy_file('figures',
                  'figures_renamed',
                  fig_file, 
                  'figure',
                  name)

print('\nFinished copying portraits, project descriptions, references and figures')
if MODIFY_FILES:
    write_errors_in_name()


# %%
