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
result = pybtex.make_bibliography('test.aux', 'plain',
                        output_backend='html'
                        )
#%%

style = pybtex.plugin.find_plugin('pybtex.style.formatting', 'unsrt')()
backend = pybtex.plugin.find_plugin('pybtex.backends', 'html')()
# parser = pybtex.database.input.bibtex.Parser()
# # with codecs.open("ref.bib", encoding="latex") as stream:
# #     # this shows what the latexcodec does to the source
# #     print(stream.read())
# with codecs.open("ref.bib", encoding="latex") as stream:
#     data = parser.parse_stream(stream)
    
# for entry in style.format_entries(data.entries.values()):
#     result = entry.text.render(backend)
# print(result)
from pybtex import auxfile

# aux_data = auxfile.parse_file('test.aux', 'bib')

result = pybtex.format_from_file('test.aux','plain',
                        output_backend='html',
                        bib_format='bibtex',
                        citations=['ref2', 'ref3', 'ref1']
                        )



# result = pybtex.database.parse_file('ref.bib')
# result.to_file('output.bib')

# result.to_file('output.html')
# document = Document()
# html_parser = HtmlToDocx()
# do stuff to document


# html_parser.add_html_to_document(result, document)

# do more stuff to document
# document.save('references.docx')
with open('result.html', 'w') as file:
    file.write(result)
#%%
    
from pybtex.plugin import find_plugin
style_cls = find_plugin('pybtex.style.formatting', 'unsrt')
style = style_cls( )
formatted_bibliography = style.format_bibliography(bib_data, ['*'])   
backend.write_to_file(formatted_bibliography, 'lol.html') 

pybtex.style.formatting.plain.Style
pybtex.style.formatting.BaseStyle

# rom pybtex.style.template
pybtex.style.formatting.unsrt
#%%

from pybtex.style.formatting.unsrt import Style as UnsrtStyle 

parser = bibtex.Parser()
bib_data = parser.parse_file('ref.bib')
entry = bib_data.entries['ref3']

his_style = UnsrtStyle()
context = {
                'entry': entry,
                'style': his_style,
                'bib_data': None,
            }

# his_style.get_article_template(entry).format_data(context)
print(his_style.get_article_template(entry).format_data(context).render(backend))
pybtex.style.names.plain.NameStyle

his_style.format_name
# join [
#     field('volume'),
#     optional ['(', field('number'),')'],
#     optional [':', pages]
# ].format_data(context)
