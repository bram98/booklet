# -*- coding: utf-8 -*-
"""
Created on Tue Jan  9 10:10:17 2024

@author: user
"""
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

os.chdir(os.path.dirname(__file__))

#%%
# Read in the data as a pandas dataframe
data = pd.read_excel('./responses.xlsx')
data.set_index('ID', inplace=True)

#%%
counts = dict()
types = dict()
for (ID, person) in data.iterrows():
    research_groups = person['Research Group'].split(';')
    for research_group in research_groups:
        if research_group == '':
            continue
        
        if not research_group in counts:
            counts[research_group] = [0,0]
        
        if not str(person['Project Type']) in types:
            types[str(person['Project Type'])] = 0
        types[str(person['Project Type'])] += 1
        
        if not person['Project Type'] is np.nan:
            counts[research_group][0] += 1
            
        counts[research_group][1] += 1
            
for (research_group, count) in counts.items():
    print(research_group[-4:-1], '\t\t', f'{count[0]}/{count[1]}' , '\t\t', f' {count[0]/count[1]*100:.3f}%' )
    