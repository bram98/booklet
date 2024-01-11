#%%
import numpy as np
import pybtex
from pybtex.database.input import bibtex
import pybtex.database.input.bibtex 
import pybtex.plugin
import codecs
import latexcodec
import os

from docx import Document
from htmldocx import HtmlToDocx


os.chdir(os.path.dirname(__file__))
#%%

from pybtex.richtext import Symbol, Text, Tag
from pybtex.style.formatting import BaseStyle, toplevel
from pybtex.style.template import (
    field, first_of, href, join, names, optional, optional_field, sentence,
    tag, together, words
)
from pybtex.style.formatting.unsrt import Style as UnsrtStyle 
class MyStyle(UnsrtStyle):
    # def format_article(self, entry_):
    #     from pybtex.textutils import abbreviate as abbr
    #     entry = entry_['entry']
        
    #     first_author = entry.persons['author'][0]
    #     author_txt = (' '.join(first_author.last_names) + ' ' +
    #                 abbr( ' '.join( first_author.first_names ) ) + ' '
    #                 )
    #     if len(entry.persons['author']) > 1:
    #         author_txt = Text(author_txt) + Tag('em', 'et al., ')
          
    #     journal_txt = ''
    #     if entry.fields['journal']:
    #         journal_txt = Text(Tag('em', entry.fields['journal'] + ' '))
            
    #     issue_txt = ''
    #     if entry.fields['journal']:
    #         journal_txt = Text(Tag('em', entry.fields['journal'] + ' '))
            
    #     return Text(author_txt,
    #                 journal_txt,
    #                 Tag('em', entry.fields['title']),)
    
#%%
parser = bibtex.Parser()
bib_data = parser.parse_file('ref.bib')
entry = bib_data.entries['ref2']

my_style = MyStyle()

formatted_bib = my_style.format_bibliography(bib_data, citations=['ref2', 'ref3'])

backend = pybtex.plugin.find_plugin('pybtex.backends', 'markdown')()

backend.write_to_file(formatted_bib, 'lol.md')

# result = backend.
# C:\Users\user\anaconda3\pkgs\pybtex-0.24.0-pyhd8ed1ab_2\site-packages\pybtex\style\formatting



