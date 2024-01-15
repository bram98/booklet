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
    data.loc[ data['Full Name'] == name, column ] = new_value

#%%
data = pd.read_excel('responses.xlsx')
data.set_index('ID', inplace=True)

data_path = Path('responses.xlsx')

columns = data.columns
print(columns)

set_entry('Kirsten Siebers', 'Photo of yourself', 'KirstenSiebers_Kirsten Siebers.jpg')




#%%

data_path.replace(data_path.with_stem('responses_old'))
data.to_excel('responses.xlsx')


#%%

