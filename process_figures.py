#%%
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from glob import glob
from pathlib import Path
import re
import os
import shutil
from tqdm import tqdm
from fuzzywuzzy import fuzz

from helper_functions import init_folder
os.chdir(os.path.dirname(__file__))

#%%

#%%
# Read in the data as a pandas dataframe
data = pd.read_excel('responses.xlsx')
data.set_index('ID', inplace=True)

#%%
def fuzzy_str_in_list(str1, strlist):
    max_ratio = 0
    for str2 in strlist:
        max_ratio = fuzz.partial_ratio(str1, str2)
        if max_ratio == 100:
            return max_ratio
    return max_ratio


def get_id(filename):
    return int(re.search('(\d+)', filename)[0])

def write_wrong_extensions_errors_and_return_valid():
    allowed_extensions = ['.jpg', '.pdf', '.ai', '.svg', '.tiff', '.psd', '.png']
    allowed_extensions_str = []
    
    extension_ids = []
    extension_errors = []
    
    figure_filenames = glob('./response_data/figures_renamed/*')
    valid_figures = []
    
    for figure in figure_filenames:
        path = Path(figure)
        ext = path.suffix
        ext = ext.lower()
        if ext in allowed_extensions:
            valid_figures.append(figure)
        else:
            ID = get_id(path.name)
            extension_ids.append(ID)
            # print(('The provided figure is not of the right file type. '
            #     f'Your file type: {ext}. Allowed file types: {" ".join(allowed_extensions)}.'
            #     ))
            extension_errors.append(('The provided figure is not of the right file type. '
                f'Your file type: {ext}. Allowed file types: {" ".join(allowed_extensions)}.'
                ))
    
    print(f'Detected {len(extension_ids)} wrong extensions')
    if MODIFY_FILES:
        write_errors('figure',extension_ids, extension_errors)
        # pass
    return valid_figures

def write_caption_errors():
    bad_people = [
        'Jonas Hehn',
        'Kristiaan Helfferich',
        'En Chen',
        'Masoud Lazemi',
        'Albaraa Falodah',
        'Hidde Nolten',
        'Kelly Brouwer',
        'Chaohong Wang',
        'Edwin Armando',
        'Ruizhi Yang',
        'Yi-Yun Lin',
        'Giovanni Del Monte'
        ]
    
    caption_ids = []
    caption_errs = []
    for(ID, person) in data.iterrows():
        if fuzzy_str_in_list( person['Full Name'], bad_people ) > 90 :
            # print(person['Full Name'], fuzzy_str_in_list( person['Full Name'], bad_people ))
            caption_ids.append(ID)
            caption_errs.append( ('Caption found inside your figure. Please remove the caption from your figure and '
                                  'add it separately as text in the "caption" option in the forms. This will allow us to '
                                  'provide a uniform style throughout the booklet.') )
    # print(len(bad_people), caption_errs[0])
    write_errors('figure', caption_ids, caption_errs)
    print(f'Detected {len(caption_ids)} captions inside figures.')



#%%

src_folder = './response_data/figures_renamed/'
dest_folder = './response_data/figures_processed/'
figure_paths = glob(src_folder + '*')

MODIFY_FILES = True 

if MODIFY_FILES:
    init_folder('figures_processed')
    

ai_counter = 0

aspect_ratios_ids = []
aspect_ratios_errs = []
print('Processing figures...')
figure_paths_valid = write_wrong_extensions_errors_and_return_valid()

figure_paths_sub = figure_paths_valid[:]

for figure_path in tqdm(figure_paths_sub):
    path = Path(figure_path)
    
    dest_path = Path(dest_folder) / path.name
    shutil.copyfile(figure_path, dest_path)
    
    # if path.suffix.lower() in ['.ai', '.psd', '.svg']:
    #     print(path.name)
    #     shutil.copyfile(
    #         dest_path, 
    #         dest_path.parent/ ( 'TEMP ' + path.stem + '.pdf' )
    #         )
    # if extension == '.ai':
    #     # convert .ai files to .jpg
    #     if MODIFY_FILES:
    #         handle_ai_file(portrait_path)
    #         portrait_path = path + '.jpg'
            
    #     ai_counter += 1
# for 
print()
if MODIFY_FILES:
    write_caption_errors()
    