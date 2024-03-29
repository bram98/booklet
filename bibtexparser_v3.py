#%%
import pybtex
from pybtex.database.input import bibtex
import pybtex.database.input.bibtex 
import pybtex.plugin
import os
import re
from importlib import reload
# from docx import Document
# from htmldocx import HtmlToDocx


os.chdir(os.path.dirname(__file__))
#%%

from pybtex.richtext import Symbol, Text
from pybtex.style.formatting import BaseStyle, toplevel
from pybtex.style.template import (
    field, first_of, href, join, names, optional, optional_field, sentence,
    tag, together, words,
    _format_data, _format_list,  Node, node, FieldIsMissing
)
from pybtex.style.formatting.unsrt import (Style as UnsrtStyle
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
    text = str(text)
    text = text.replace(u"\u2013", "-")
    text = text.replace("--", "-")
    # print(text, Text(r'\\-\\-').join(text.split(dash_re)))
    # print(dash_re.search(str(text)), str(text))
    if dash_re.search(str(text)):
        return Text( 'pp. ', Text('--').join(text.split('-')), '' )
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
    formatted_names = [style.my_format_name(person, abbr=True) for person in persons[:max_names]]

    formatted_names = join(**kwargs) [ formatted_names ]
    if len(persons) > max_names or True:
        # always add et al., since in the forms many people gave just one name, while there are usually more than 1 authors.
        formatted_names = join [ formatted_names, ' ', tag('em') [ 'et al.' ] ]
        
    return formatted_names.format_data(context)

class MyStyle(UnsrtStyle):
    abbreviate_names = True
    def my_format_name(self, person, abbr=True):
        # print(abbr)
        return join[
            name_part(tie=True, abbr=abbr)[person.rich_first_names + person.rich_middle_names],
            name_part(tie=True)[person.rich_prelast_names],
            name_part[person.rich_last_names],
            name_part(before=', ')[person.rich_lineage_names]
        ]
            
    def format_names(self, role, as_sentence=True):
        formatted_names = short_names(role, max_names=1, sep=', ', sep2 = ' and ', last_sep=', and ')

        if as_sentence:
            return sentence [ formatted_names ]
        else:
            return formatted_names
    def get_inproceedings_template(self, e):
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

        template = toplevel [
            join(', ') [
                self.format_names('author'),
                # optional [ '\"', field('title'), '\"'],
                sentence [
                    tag('em') [ strict_first_of[ field('journal'), field('organization') ] ],
                    join(', ') [ 
                        optional [ volume_and_issue_number ],
                        optional [ pages ],
                        optional [ field('year') ],
                    ],
                ],
            ]

        ]
        return template
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

        template = toplevel [
            join(', ') [
                self.format_names('author'),
                # optional [ '\"', field('title'), '\"'],
                sentence [
                    tag('em') [ field('journal') ],
                    join(', ') [ 
                        optional [ volume_and_issue_number ],
                        optional [ pages ],
                        optional [ field('year') ],
                    ],
                ],
            ]

        ]
        return template
           # Tag('em', entry.fields['title']),)
           
def make_bibliography(bib_file, output_filename, output_file_type='markdown', citations=['*']):
    if not output_file_type in ['html', 'markdown']:
        raise Exception(f'output_file_type not html or markdown, but {output_file_type}')
        
    parser = bibtex.Parser()
    bib_data = parser.parse_file(bib_file)
    # entry = bib_data.entries['ref2']

    my_style = MyStyle()

    formatted_bib = my_style.format_bibliography(bib_data, citations=citations)

    backend = pybtex.plugin.find_plugin('pybtex.backends', output_file_type)()

    output_extension = {
        'html': '.html',
        'markdown': '.md'
        }[output_file_type]
    backend.write_to_file(formatted_bib, output_filename + output_extension)
#%%

