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
parser = bibtex.Parser()
bib_data = parser.parse_file('ref.bib')

#%%
ref1 = bib_data.entries['ref1']
print(ref1.items())
#%%
os.system('pdflatex test.tex')
#%%

# style = pybtex.plugin.find_plugin('pybtex.style.formatting', 'plain')()
# backend = pybtex.plugin.find_plugin('pybtex.backends', 'html')()
# parser = pybtex.database.input.bibtex.Parser()
# # with codecs.open("ref.bib", encoding="latex") as stream:
# #     # this shows what the latexcodec does to the source
# #     print(stream.read())
# with codecs.open("ref.bib", encoding="latex") as stream:
#     data = parser.parse_stream(stream)
    
# for entry in style.format_entries(data.entries.values()):
#     result = entry.text.render(backend)
# print(result)

# result = pybtex.format_from_file('ref.bib','plain',
#                         output_backend='html',
#                         bib_format='bibtex', 
#                         output_encoding='html'
#                         )

result = pybtex.make_bibliography('ref.bib','plain',
                        output_backend='html',
                        bib_format='bibtex'
                        )
document = Document()
html_parser = HtmlToDocx()
# do stuff to document


html_parser.add_html_to_document(result, document)

# do more stuff to document
# document.save('references.docx')
with open('result.html', 'w') as file:
    file.write(result)