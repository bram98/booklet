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
from pathlib import Path
os.chdir(os.path.dirname(__file__))

from helper_functions import parse_id, write_errors

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
    Checks whether an image has the right aspect ratio of w x h == 35 x 45. 
    Returns false and an error message for an incorrect aspect ratio and returns
    True for a valid aspect ratio.
    '''
    aspect_ratio = img.size[0]/img.size[1]

    ideal_aspect_ratio = 35/45
    if abs(aspect_ratio - ideal_aspect_ratio) > 0.001:
        return ( False, (f'Your portrait photo has the wrong aspect ratio. '
                f'Required:   w : h = 35 : 45. Your photo:   {img.size[0]} : {img.size[1]}. '
                f'Please crop your photo until it is 35 by 45 mm, or something with the same aspect ratio.'
               ) )
    return ( True, '')

def check_resolution(img: wImg) -> (bool, str):
    '''
    Checks whether image is of high enough resolution. Uses a cutoff of 190 horizontal
    pixels, determined by eye. Returns false and an error message for low resolution
    and returns True for a valid resolution.
    '''
    
    if img.size[0] < 190:
        img.units = 'pixelsperinch'
        dpi = int(np.round(25.4/35*img.size[0])) 
        return ( False, (f'It has been detected that your portrait photo has low resolution. '
                         f'It is recommended to have at least 140 dpi, or to be '
                         f'193 pixels wide. Your photo has {dpi} dpi and is {img.size[0]} '
                         f'pixels wide. Please upload a photo with higher resolution.'
                         ))
    return (True, '')


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
portrait_files = glob(src_folder + '*')

MODIFY_FILES = True 

portraits = glob(src_folder + '*')
if MODIFY_FILES:
    init_folder('portraits_processed')
    for portrait_file in portrait_files:
        src_path = Path(portrait_file)
        target_path = Path(dest_folder)/src_path.name
        shutil.copy(src_path, target_path)

#%%
MODIFY_FILES = True

if MODIFY_FILES:
    ai_files = glob(dest_folder + '*.ai')
    print('Converting .ai files...')
    for ai_file in tqdm(ai_files):
        handle_ai_file(ai_file)
    print()
    print(f'Done converting {len(ai_files)} .ai files')
    
    png_files = glob(dest_folder + '*.png')
    
    print('Converting .png files...')
    for png_file in tqdm(png_files):
        with wImg(filename=png_file, resolution=300) as img:
            img.compression_quality = 99
            jpg_file = Path(png_file).with_suffix('.jpg')
            # print(jpg_file)
            img.save(filename=jpg_file)
        Path(png_file).unlink(missing_ok=True)
    print()
    print(f'Done converting {len(png_files)} .png files to .jpg')
#%%
src_folder = './response_data/portraits_renamed/'
dest_folder = './response_data/portraits_processed/'
portrait_paths = glob(src_folder + '*')

MODIFY_FILES = True 

if MODIFY_FILES:
    init_folder('portraits_processed')

ai_counter = 0
portrait_paths_sub = portrait_paths[:20]

resolution_ids = []
resolution_errs = []
aspect_ratios_ids = []
aspect_ratios_errs = []
print('Processing portraits...')
for portrait_path in tqdm(portrait_paths):
# for portrait_path in portrait_paths:
    path, extension = os.path.splitext(portrait_path)
    
    if extension == '.ai':
        # convert .ai files to .jpg
        if MODIFY_FILES:
            handle_ai_file(portrait_path)
            portrait_path = path + '.jpg'
            
        ai_counter += 1
    # try:
    #     with wImg(filename=portrait_path) as img:
    #         portrait_file = os.path.split(portrait_path)[1]
            
    #         (valid_resolution, err_msg) = check_resolution(img)
    #         if not valid_resolution:
    #             resolution_ids.append(parse_id(portrait_file))
    #             resolution_errs.append(err_msg)
           
    #         (valid_aspect_ratio, err_msg ) = check_aspect_ratio(img)
    #         if not valid_aspect_ratio:
    #             aspect_ratios_ids.append(parse_id(portrait_file))
    #             aspect_ratios_errs.append(err_msg)
            
    #         if (not valid_aspect_ratio) or (not valid_resolution):
    #             continue
            
    #         if MODIFY_FILES:
                
    #             rescale_dpi(img)
    #             if img.format == 'png':
    #                 with img.convert('jpg', resolution=300) as converted:
    #                     converted.save(filename='converted.png')
    #                 print('png', path)
    #             img.save(filename=dest_folder + portrait_file)
    #             # 
    #         # print(np.array(img.size)/np.array(img.resolution)*25.4)
            
    try:
        with wImg(filename=portrait_path) as img:
            target_path = Path(dest_folder)/Path(portrait_path).name
            portrait_file = target_path.name
            # print( img.format )
            (valid_resolution, err_msg) = check_resolution(img)
            if not valid_resolution:
                resolution_ids.append(parse_id(portrait_file))
                resolution_errs.append(err_msg)
           
            (valid_aspect_ratio, err_msg ) = check_aspect_ratio(img)
            if not valid_aspect_ratio:
                aspect_ratios_ids.append(parse_id(portrait_file))
                aspect_ratios_errs.append(err_msg)
            
            if (not valid_aspect_ratio) or (not valid_resolution):
                continue
            
            if MODIFY_FILES:
                
                rescale_dpi(img)
                # print(img.format)
                if img.format == 'PNG':
                    with img.convert('jpg', resolution=300) as converted:
                        # converted.save(filename='converted.png')
                        # img.format = 'JPG'
                        target_path = target_path.with_suffix('.jpg')
                        
                        converted.save(filename=target_path)
                else:
                    img.save(filename=target_path)
                # 
            # print(np.array(img.size)/np.array(img.resolution)*25.4)
            
            
    except Exception as e:
        print(f'Error trying to open {portrait_path}')
        raise e

print()  
if MODIFY_FILES:
    write_errors('portrait', resolution_ids, resolution_errs)
    write_errors('portrait', aspect_ratios_ids, aspect_ratios_errs)

print('\nDone processing portraits...')
print( f'Detected {len(resolution_ids)} photos with low resolution.')
print( f'Detected {len(aspect_ratios_ids)} wrong portrait aspect ratios.')
print(f'Converted {ai_counter} .ai file(s) to .jpg')
# %%
