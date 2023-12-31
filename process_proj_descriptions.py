#%%
import numpy as np
import matplotlib.pyplot as plt
import os
from pathlib import Path
from glob import glob
from docx import Document
from docx.shared import Inches
import doc2docx
from tqdm import tqdm
import re

#%%

os.chdir(os.path.dirname(__file__))

#%%
def convert_doc_to_docx_windows():
    print('Converting .doc files to .docx files...')
    doc_files = glob('response_data/project_descriptions_renamed/*.doc')
    for doc_file in tqdm(doc_files):
        doc2docx.convert(doc_file)
    print('\nDone converting .doc files to .docx files')
    
    print('Removing old .doc files...')
    for doc_file in tqdm(doc_files):
        Path(doc_file).unlink( missing_ok=True)
    
    print('\nDone removing old .doc files...')

def inspect_file_types():
    filenames = glob('response_data/project_descriptions_renamed/*')
    for filename in filenames:
        ext  = Path(filename).suffix
        if ext != '.docx':
            print(Path(filename).name)

def write_file_type_errors():
    filenames = glob('response_data/project_descriptions_renamed/*')
    filename_ids = []
    filename_errs = []
    for filename in filenames:
        file_path = Path(filename)
        ext = file_path.suffix
        if ext != '.docx':
            filename_ids.append( int( parse_id( file_path.name ) ) )
            filename_errs.append( (f'The file type of your project description '
                                   f'is not supported. Your file type: {ext}. '
                                   f'Supported: .doc or .docx.'
                                   ))
    print(f'Detected {len(filename_errs)} wrong projection description file types.')
    write_errors('proj_description', filename_ids, filename_errs)
    # write_errors('figure', caption_ids, caption_errs)
# convert_doc_to_docx_windows()
# inspect_file_types()

#%%
convert_doc_to_docx_windows()
write_file_type_errors()
#%%
filenames = glob('response_data/project_descriptions_renamed/*.docx')

reference_regex = re.compile(r'(^\[\d\])|references:?', flags=re.M|re.I)
info_regex = re.compile(r'^(sponsors?)|(supervisors?)|(keywords)', flags=re.M|re.I)
figure_regex = re.compile(r'^fig(ure)?', flags=re.I)
email_regex = re.compile(r'@', flags=re.I)
# lengths = []
for filename in filenames[:40]:
    doc = Document(filename)
    # print( [len(p.text) for p in doc.paragraphs] )
    paras = [p.text for p in doc.paragraphs]
    for para in paras:
        if len(para) == 0:
            continue
        if not reference_regex.search(para) is None:
            # print(para)
            # print(len(para))
            break
        if not info_regex.search(para) is None:
            break
        
        if not figure_regex.search(para) is None:
            # print(para)
            break
        if not email_regex.search(para) is None:
            print(para)
            # break
        
        if len(para) < 250:
            # print(para)
            print('')
    # lengths.append([len(p.text) for p in doc.paragraphs])
    

# lengths_flattened = [length for row in lengths for length in row]

#%%
plt.figure()
plt.ylim(0,50)
plt.hist(lengths_flattened, bins=200)
