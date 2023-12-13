#%%
from glob import glob
from PIL import Image
import os
import shutil
from wand.image import Image as wImg
from wand.display import display
import cv2 as cv
import matplotlib.pyplot as plt
from pdf2image import convert_from_path as pdf_read
from tqdm import tqdm
import re
import numpy as np

#%%
def handle_ai_file(file_path:str):
    # folder, filename_full = os.path.split(file_path)
    filename, extension = os.path.splitext(file_path)
    file_path_pdf = filename + '.pdf'
    os.rename(file_path, file_path_pdf)
    
    pdf_file = pdf_read(file_path_pdf, 400)[0]
    pdf_file.save(f'{filename}.jpg')
    os.remove(file_path_pdf)
    
def init_folder(folder:str):
    '''
    Creates folder if it does not exist and if it exists, it will delete all files inside.
    '''
    os.makedirs(f'./response_data/{folder}',exist_ok=True)
    print(f'Created folder {folder}')
    files = glob(f'./response_data/{folder}/*')
    for file in files:
        os.remove(file)

def check_aspect_ratio(img: wImg) -> (bool, str):
    '''
    Checks whether an image has the right aspect ratio of w x h == 35 x 45
    '''
    aspect_ratio = img.size[0]/img.size[1]

    ideal_aspect_ratio = 35/45
    if abs(aspect_ratio - ideal_aspect_ratio) > 0.001:
        return ( False, (f'Your portrait photo has the wrong aspect ratio. '
                f'Required:   w : h = 35 : 45. Your photo:   {img.size[0]} : {img.size[1]}'
               ) )
    return ( True, '')

def parse_id(filename: str) -> int:
    '''
    Return id from filename. Assumes the ID are the only numbers in the string.
    '''
    only_numbers = re.compile(r'\d+')
    return int(only_numbers.search(filename)[0])

def rescale_dpi(img: wImg):
    # if img.units != 'pixelsperinch':
    #     raise Exception(f'Image not right units: {img.units}')
    
    img.units = 'pixelsperinch'
    # The dpi should be the same x and y if the aspect ratio is the same.
    dpi = int(np.round(25.4/35*img.size[0])) 
    img.resolution = (dpi, dpi)
#%%
src_folder = './response_data/portraits_renamed/'
dest_folder = './response_data/portraits_processed/'
portrait_paths = glob(src_folder + '*')

MODIFY_FILES = True 

if MODIFY_FILES:
    init_folder('portraits_processed')

ai_counter = 0
portrait_paths_sub = portrait_paths[:20]

aspect_ratios_ids = []
aspect_ratios_errs = []
print('Processing portraits...')
for portrait_path in tqdm(portrait_paths):
    path, extension = os.path.splitext(portrait_path)
    

    if extension == '.ai':
        # convert .ai files to .jpg
        if MODIFY_FILES:
            handle_ai_file(portrait_path)
            portrait_path = path + '.jpg'
            
        ai_counter += 1
    try:
        with wImg(filename=portrait_path) as img:
            portrait_file = os.path.split(portrait_path)[1]
            
            (valid_aspect_ratio, err_msg ) = check_aspect_ratio(img)
            if not valid_aspect_ratio:
                aspect_ratios_ids.append(parse_id(portrait_file))
                aspect_ratios_errs.append(err_msg)
                continue
            
            if MODIFY_FILES:
                rescale_dpi(img)
                if img.format == 'png'
                img.save(filename=dest_folder + portrait_file)
            # print(np.array(img.size)/np.array(img.resolution)*25.4)
            
    except Exception as e:
        print(f'Error trying to open {portrait_path}')
        raise e
write_errors('portrait', aspect_ratios_ids, aspect_ratios_errs)

print('\nDone processing portraits...')
print( f'Detected {len(aspect_ratios_ids)} wrong portrait aspect ratios.')
print(f'Converted {ai_counter} .ai file(s) to .jpg')
# %%
