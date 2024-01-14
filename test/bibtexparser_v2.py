#%%
import pybtex
from pybtex.database.input import bibtex
import pybtex.database.input.bibtex 
import pybtex.plugin
import os
import re

from docx import Document
from htmldocx import HtmlToDocx


os.chdir(os.path.dirname(__file__))
#%%

from pybtex.richtext import Symbol, Text, Tag
from pybtex.style.formatting import BaseStyle, toplevel
from pybtex.style.template import (
    field, first_of, href, join, names, optional, optional_field, sentence,
    tag, together, words,
    _format_data, _format_list,  Node, node, FieldIsMissing
)
from pybtex.style.formatting.unsrt import (Style as UnsrtStyle,
                                           dashify, date
                                           )
from pybtex.style.names import BaseNameStyle, name_part


@node
def strict_first_of(children, context):
    """Return first nonempty child. If no entry is found, raise error"""
    
    for child_node in children:
        child = optional [ child_node ]
        formatted_child = _format_data(child, context)
        if formatted_child:
            return formatted_child
    
    raise FieldIsMissing('', context['entry'])

def pagify(text):
    dash_re = re.compile(r'[-\-]+')
    print(dash_re.search(str(text)), str(text))
    if dash_re.search(str(text)):
        return Text( 'pp. ', Symbol('ndash').join(text.split(dash_re)), '' )
    else:
        return Text( 'p. ', text )
    
def dashify(text):
    dash_re = re.compile(r'-+')
    return Text(Symbol('ndash')).join(text.split(dash_re))

pages = optional [ 
        strict_first_of [ 
            field('pages', apply_func=pagify) ,
            field('page', apply_func=pagify),
            field('page range', apply_func=pagify) 
        ]
    ]


@node
def short_names(children, context, role, max_names=1, **kwargs):
    """Return formatted names."""

    assert not children

    try:
        persons = context['entry'].persons[role]
    except KeyError:
        raise FieldIsMissing(role, context['entry'])

    style = context['style']
    # print(persons[:max_names])
    formatted_name = [style.my_format_name(person, abbr=True) for person in persons[:max_names]]
    # formatted_names = style.my_format_name(persons[0], abbr=True)
    print(formatted_name)
    formatted_name = join(**kwargs) [formatted_name]
    if len(persons) > max_names:
        formatted_name = join [ formatted_name, tag('em') [ ' et al.' ] ]
    print(formatted_name) 
    # formatted_name = [style.format_name(person, True) for person in persons]
    # return join(**kwargs) [formatted_name].format_data(context)
    return formatted_name.format_data(context)

class MyStyle(UnsrtStyle):
    abbreviate_names = True
    def my_format_name(self, person, abbr=True):
        print(abbr)
        return join[
            name_part(tie=True, abbr=abbr)[person.rich_first_names + person.rich_middle_names],
            name_part(tie=True)[person.rich_prelast_names],
            name_part[person.rich_last_names],
            name_part(before=', ')[person.rich_lineage_names]
        ]
            
    def format_names(self, role, as_sentence=True):
        formatted_names = short_names(role, max_names=3, sep=', ', sep2 = ' and ', last_sep=', and ')
        # print(formatted_names.format_data(context))
        # print(len(formatted_names))
        # print(formatted_names)
        if as_sentence:
            return sentence [ formatted_names ]
        else:
            return formatted_names
    
    def get_article_template(self, e):
        volume_and_issue_number = (# volume and pages, with optional issue number
            join [
                optional [ tag('em') [ field('volume') ] ],
                optional [ 
                    '(',
                    strict_first_of [ 
                        field('issue'), 
                        field('issues'), 
                        field('number'),
                    ],
                    ')',
                ],
            ],
        )
        # print(optional [ field('pages') ].format_data(context))
        # volume_and_issue_number = '3'
        # print(self.format_names('author'))
        template = toplevel [
            join(', ') [
                self.format_names('author'),
                # optional [ '\"', field('title'), '\"'],
                sentence [
                    tag('em') [ field('journal') ],
                    join(', ') [ 
                        optional [ volume_and_issue_number ],
                        optional [ pages ],
                    ],
                ],
                ]
            # optional [ '(', field('year'), ')' ],
            
            # sentence [ optional_field('note') ],
            # self.format_web_refs(e),
        ]
        return template
           # Tag('em', entry.fields['title']),)
           
my_style = MyStyle()    
formatted_bib = my_style.format_bibliography(bib_data, citations=['ref2', 'ref1', 'ref3'])

backend = pybtex.plugin.find_plugin('pybtex.backends', 'markdown')()
backend_html = pybtex.plugin.find_plugin('pybtex.backends', 'html')()

backend.write_to_file(formatted_bib, 'lol.md')
backend_html.write_to_file(formatted_bib, 'lol.html')
#%%
parser = bibtex.Parser()
bib_data = parser.parse_file('ref.bib')
entry = bib_data.entries['ref2']

my_style = MyStyle()

formatted_bib = my_style.format_bibliography(bib_data, citations=['ref2', 'ref1', 'ref3', 'ref4'])

backend = pybtex.plugin.find_plugin('pybtex.backends', 'markdown')()
backend_html = pybtex.plugin.find_plugin('pybtex.backends', 'html')()

backend.write_to_file(formatted_bib, 'lol.md')
backend_html.write_to_file(formatted_bib, 'lol.html')

# result = backend.
# C:\Users\user\anaconda3\pkgs\pybtex-0.24.0-pyhd8ed1ab_2\site-packages\pybtex\style\formatting



#%%
my_style.format_name