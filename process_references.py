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
src_folder = './response_data/references_renamed/'
reference_paths = glob(src_folder + '*')

for reference_path in reference_paths:
    ID = parse_id(reference_path)
    write_errors_references(
            ID,
            reference_path
        )
# for(ID, person) in tqdm(data.iterrows(), total=len(data)):
#     write_errors_references(
#         ID,
#         f'./response_data/project_descriptions/{ref_file}'
#     )