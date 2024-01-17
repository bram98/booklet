#%%
import pandas as pd
from os import makedirs
from shutil import copy
import os.path as osp
from glob import glob
import pypandoc as pypandoc
import numpy as np
from unidecode import unidecode
from pandas.api.types import is_any_real_numeric_dtype
import os
import re

from docx import Document
from docx.shared import Pt

os.chdir(os.path.dirname(__file__))
#%%
df = pd.read_excel('responses.xlsx')


first = f''
last = f''


group_regex = re.compile('\((.+?)\)')
# For each response, collect everything in a folder for Conny
# for i, row in df.head(2).iterrows(): 
for i, row in df.iterrows():
    try:
        row_ =row
        first_group = row['Research Group'].split(';')[0].split('(')[0].strip()
        group_abbr = group_regex.search(row["Research Group"]).group(1)
        if not 'NP' in group_abbr:
            continue
        full_name = row['Full Name'].strip().split()
        id_and_name = f'{row["ID"]} {unidecode(row["Full Name"])}'
        print(f'{group_abbr:4s}\t {id_and_name}')
        dest_dir = f'responses-conny/{first_group}/{full_name[-1]}, {" ".join(full_name[:(-1)])}'
        makedirs(dest_dir, exist_ok=True)
    
        # Copy the image files
        file_src = glob(f'response_data/portraits_renamed/portrait {id_and_name}.*')[0]
        file_dest = osp.join(dest_dir, osp.basename(file_src))
        copy(file_src, file_dest)
    
        file_src = glob(f'response_data/figures_renamed/figure {id_and_name}.*')
        if len(file_src) > 0:  # It's allowed for responses to have no figures
            file_src = file_src[0]
            file_dest = osp.join(dest_dir, osp.basename(file_src))
            copy(file_src, file_dest)
            
        tab = '&Tab;'
        # Create the word file with project description and other info
        md_string = first
        if row['Daily Supervisor(s) (optional)'] is np.nan or row['Daily Supervisor(s) (optional)'] == row['Supervisor(s)']:
            daily_str =''
        else:
            daily_str = f'\n| Daily supervisor(s):{tab}{row["Daily Supervisor(s) (optional)"]}\n'
        md_string += f"""\
{row['Project Title']}

| {row['Full Name']} ({row['Project Type']}), {row['E-mail']}

| Sponsor:{tab*2}{row['Sponsor(s)']}
| Supervisor(s):{tab}{row['Supervisor(s)']}{daily_str}
| Keyword(s):{tab*2}{row['Keywords']}

"""
        file_src = glob(f'response_data/project_descriptions_processed/project_description {id_and_name}.*')[0]
        # project_description = pandoc.read(file=file_src)
        project_description = pypandoc.convert_file(source_file=file_src, to='md')
        # md_string += pandoc.write(project_description)
        md_string += project_description
        
        md_string += '\n\n'
        
        file_src = glob(f'response_data/references_processed/references {id_and_name}.md')
        emdash= u'\u2014'
        endash = u"\u2013"
        if len(file_src) == 1:
            file_src = file_src[0]
            with open(file_src, 'r') as file:
                ref_string = file.read()
                
                print(ref_string)
                ref_string = ref_string.replace('\-\-', endash)
                md_string += ref_string
                print(ref_string)
        
        # print(md_string)
        fig_caption = row['Figure Caption (optional)']
        if not (is_any_real_numeric_dtype(type(fig_caption)) and np.isnan(fig_caption)):
            fig_caption = fig_caption.removeprefix('Figure 1').removeprefix('figure 1').removeprefix(':').removeprefix('.')
            md_string += f"""
    
Figure 1: {fig_caption.strip()}
    
    """
    
        # Here we have to add the references (bibliography) to the md_string
        
        # Write the md_string to a Word file
        # doc = pandoc.read(md_string)
        # file_dest = osp.join(dest_dir, f'project_description {row["ID"]} {row["Full Name"]}.docx')
        outputfile =  osp.join(dest_dir, f'project_description {row["ID"]} {row["Full Name"]}.docx')
        # print(md_string)
        pypandoc.convert_text(md_string, format='md', to='docx', outputfile=outputfile)
        
        doc = Document(outputfile)
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Arial'
        font.size = Pt(12)
        
        for para in doc.paragraphs:
            para.style = doc.styles['Normal']
            
        doc.save(outputfile)
        
    except Exception as e:
        print(repr(e))
        raise
        