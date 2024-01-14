#%%
import numpy as np
import pandas as pd
# import PIL
from  glob import glob
import os
from os.path import splitext
# import shutil
from unidecode import unidecode
from urllib.parse import urlparse, parse_qs, unquote
import re
from tqdm import tqdm
from pandas.api.types import is_integer_dtype
from pandas.api.types import is_any_real_numeric_dtype
from pathlib import Path
from docx import Document

from helper_functions import init_folder, parse_id, copy_file
from bibtexparser_v3 import make_bibliography

os.chdir(os.path.dirname(__file__))
#%%
# Read in the data as a pandas dataframe
data = pd.read_excel('responses.xlsx')
data.set_index('ID', inplace=True)
#%%

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
        err_list.append('Reference error. The first column with reference numbers does not contain (only) numbers.')
    if not is_integer_dtype(df.iloc[:,3]):
        id_list.append(person_id)
        err_list.append('Reference error. The fourth column with journal volume/issue does not contain (only) numbers.')
    if not is_integer_dtype(df.iloc[:,5]):
        id_list.append(person_id)
        err_list.append('Reference error. The sixth column with publication year does not contain (only) numbers.')
    for i in range(df.columns.size):
        if np.any(df.iloc[:,i].isna()):
            id_list.append(person_id)
            err_list.append(f'Reference error. Column {i+1} contains some empty cells')

    write_errors('references', id_list, err_list)
#%%



def check_page_range(pr):
    if is_any_real_numeric_dtype(type(pr)) and np.isnan(pr):
        return False
    else:
        return True


def xlsx2bib(xlsx_file):
    # Replace file extension
    base_name = splitext(xlsx_file)[0]
    bib_file = base_name + '.bib'

    # Parse excel file
    df = pd.read_excel(xlsx_file).sort_values(by='Number')

    # For each row in the xlsx file, generated an entry in the bib file
    with open(bib_file, 'w') as fh:
        for i, row in df.iterrows():
            fh.write(f"""@article{{ref{row['Number']},
            author={{{row['First author'].replace('&', chr(92)+'&').strip()}}},
            journal={{{row['Journal'].strip()}}},
            number={{{int(row['Issue']) if not np.isnan(row['Issue']) else ''}}},
            pages={{{str(row['page range']).replace('-', '--').strip() if check_page_range(row['page range']) else ''}}},
            year={{{row['year']}}},
            }}
            """)
#%%
src_folder = './response_data/references_renamed/'
dest_folder = './response_data/references_processed/'
reference_paths = glob(src_folder + '*')

MODIFY_FILES = True

if MODIFY_FILES:
    init_folder('references_processed')
    
'''
Convert excel to .bib
'''
for reference_path in tqdm(reference_paths[:]):
    ref_path = Path(reference_path)
    ref_name = ref_path.stem
    ref_path = copy_file(
        src_folder='references_renamed', 
        dest_folder='references_processed',
        old_filename=ref_path.name,
        new_filename='references',
        name = ref_name[11:]
        )
    ref_path = Path(ref_path)
    
    try:
        xlsx2bib(ref_path.absolute())
    except Exception as e:
        print(f'[{ref_name[11:]}] {repr(e)}')
        
#%%
bib_paths = glob(dest_folder + '*.bib')
proj_description_paths = glob('./response_data/project_descriptions_renamed/*')


cite_sb_regex = re.compile(r'(\[.*?\d.*?\])') # Matches [1] and [ 1 ]
cite_sb_regex2 = re.compile(r'\[.*?(\d+)\D*?(\d+)?\]') # Matches [1], [1, 2] and [8-10]
cite_cite_regex = re.compile(r'\\cite')
for bib_path in bib_paths:
    bib_path = Path(bib_path)
    person = bib_path.stem[11:]

    proj_descriptions = list(filter(lambda file: person in file, proj_description_paths ))
    
    if len(proj_descriptions) != 1:
        raise Exception(f'[{person}] Found {len(proj_descriptions)} project descriptions!')
        
    proj_description = proj_descriptions[0]
    
    doc = Document(proj_description)
    # print( [len(p.text) for p in doc.paragraphs] )
    paras = ''.join( [p.text for p in doc.paragraphs] )
    
    # Find all references and then select the largest one.
    numbers = cite_sb_regex2.findall(paras)
    numbers = np.array(numbers, dtype=str).flatten()
    numbers = np.array(list(filter(lambda x: x, numbers)), dtype=int)
    if len(numbers) == 0:
        print(person)
    else:
        largest_ref = np.max(numbers)



