from  glob import glob
import os
import shutil
import json
from fuzzywuzzy import fuzz
import re

def init_folder(folder):
    '''
    Creates folder if it does not exist and if it exists, it will delete all files inside.
    '''
    os.makedirs(f'./response_data/{folder}',exist_ok=True)
    print(f'Created folder {folder}')
    files = glob(f'./response_data/{folder}/*')
    for file in files:
        os.remove(file)
        
def copy_file(src_folder, dest_folder, old_filename, new_filename, name):
    '''
    Copies a file from src_folder/old_filename to dest_folder/new_filename {name}.oldextension
    '''
    try:
        filename, file_extension = os.path.splitext(old_filename)
        # if file_extension == '.ai':
        #     file_extension = '.jpg' # I manually converted .ai to .jpg
        src_file_name = f'./response_data/{src_folder}/{filename}{file_extension}'
        dest_file_name = f'./response_data/{dest_folder}/{new_filename} {name}{file_extension}'
        
        shutil.copy(src_file_name, dest_file_name)
        return dest_file_name
    except Exception as e:
        print(f'Error! Name:{name}\t File name:{old_filename}')
        raise e
        
def fuzzy_equal(str1, str2):
    return fuzz.ratio(str1, str2) > 80

def write_errors(err_type, id_list, err_list):
    """Writes errors in data to error_data.json. This helps in creating an e-mail request for improving responses.

    Parameters
    ----------
    err_type: str
        error type. Must be one of 'portrait', 'figure', proj_description', 'references' or 'name'
    id_list: list(int)
        list of all people with errors specified by their ID
    err_list: list(str)
        list of strings that will appear in literally in the final email. Same order as id_list
    
    Returns
    -------
        None
    """
    if len(id_list)==0:
        return
    if not err_type in ['portrait', 'proj_description', 'figure', 'references', 'name']:
        raise Exception(("Error type must be one of 'portrait', 'proj_description', 'figure', 'references' or 'name'. "
                         f"Not '{err_type}'"
                         ))
    
    with open('error_data.json') as file:
        error_dict = json.load(file)
        for (i, person_id) in enumerate(id_list):
            
            # list of errors for this person currently in error_dict
            person_err_list = error_dict[str(person_id)][f'{err_type}_errors']
            
            # checks if error is already present (prevents writing same error twice)
            if not any( fuzzy_equal(err_list[i], old_err) for old_err in person_err_list ):
                person_err_list.append(err_list[i])
    
    with open('error_data.json', 'w') as file:
        json.dump(error_dict, file, indent=4)
    
    print(f'Wrote {len(err_list)} error(s) of type {err_type} to error_data.json')

def clear_errors(err_type: str):
    if not err_type in ['portrait', 'proj_description', 'figure', 'references', 'name']:
        raise Exception(("Error type must be one of 'portrait', 'proj_description', 'figure', 'references' or 'name'. "
                         f"Not '{err_type}'"
                         ))
        
    with open('error_data.json') as file:
        error_dict = json.load(file)

    for key in error_dict.keys():
        error_dict[key][f'{err_type}_errors'] = []
    
    with open('error_data.json', 'w') as file:
        json.dump(error_dict, file, indent=4)
        
def parse_id(filename: str) -> int:
    '''
    Return id from filename. Assumes the ID are the only numbers in the string.
    '''
    only_numbers = re.compile(r'\d+')
    return int(only_numbers.search(filename)[0])

latexAccents = [
  [ u"à", "\\`a" ], # Grave accent
  [ u"è", "\\`e" ],
  [ u"ì", "\\`\\i" ],
  [ u"ò", "\\`o" ],
  [ u"ù", "\\`u" ],
  [ u"ỳ", "\\`y" ],
  [ u"À", "\\`A" ],
  [ u"È", "\\`E" ],
  [ u"Ì", "\\`\\I" ],
  [ u"Ò", "\\`O" ],
  [ u"Ù", "\\`U" ],
  [ u"Ỳ", "\\`Y" ],
  [ u"á", "\\'a" ], # Acute accent
  [ u"é", "\\'e" ],
  [ u"í", "\\'\\i" ],
  [ u"ó", "\\'o" ],
  [ u"ú", "\\'u" ],
  [ u"ý", "\\'y" ],
  [ u"Á", "\\'A" ],
  [ u"É", "\\'E" ],
  [ u"Í", "\\'\\I" ],
  [ u"Ó", "\\'O" ],
  [ u"Ú", "\\'U" ],
  [ u"Ý", "\\'Y" ],
  [ u"â", "\\^a" ], # Circumflex
  [ u"ê", "\\^e" ],
  [ u"î", "\\^\\i" ],
  [ u"ô", "\\^o" ],
  [ u"û", "\\^u" ],
  [ u"ŷ", "\\^y" ],
  [ u"Â", "\\^A" ],
  [ u"Ê", "\\^E" ],
  [ u"Î", "\\^\\I" ],
  [ u"Ô", "\\^O" ],
  [ u"Û", "\\^U" ],
  [ u"Ŷ", "\\^Y" ],
  [ u"ä", "\\\"a" ],    # Umlaut or dieresis
  [ u"ë", "\\\"e" ],
  [ u"ï", "\\\"\\i" ],
  [ u"ö", "\\\"o" ],
  [ u"ü", "\\\"u" ],
  [ u"ÿ", "\\\"y" ],
  [ u"Ä", "\\\"A" ],
  [ u"Ë", "\\\"E" ],
  [ u"Ï", "\\\"\\I" ],
  [ u"Ö", "\\\"O" ],
  [ u"Ü", "\\\"U" ],
  [ u"Ÿ", "\\\"Y" ],
  [ u"ç", "\\c{c}" ],   # Cedilla
  [ u"Ç", "\\c{C}" ],
  [ u"œ", "{\\oe}" ],   # Ligatures
  [ u"Œ", "{\\OE}" ],
  [ u"æ", "{\\ae}" ],
  [ u"Æ", "{\\AE}" ],
  [ u"å", "{\\aa}" ],
  [ u"Å", "{\\AA}" ],
  [ u"–", "--" ],   # Dashes
  [ u"—", "---" ],
  [ u"ø", "{\\o}" ],    # Misc latin-1 letters
  [ u"Ø", "{\\O}" ],
  [ u"ß", "{\\ss}" ],
  [ u"¡", "{!`}" ],
  [ u"¿", "{?`}" ],
  [ u"\\", "\\\\" ],    # Characters that should be quoted
  [ u"~", "\\~" ],
  [ u"&", "\\&" ],
  [ u"$", "\\$" ],
  [ u"{", "\\{" ],
  [ u"}", "\\}" ],
  [ u"%", "\\%" ],
  [ u"#", "\\#" ],
  [ u"_", "\\_" ],
  [ u"≥", "$\\ge$" ],   # Math operators
  [ u"≤", "$\\le$" ],
  [ u"≠", "$\\neq$" ],
  [ u"©", "\copyright" ], # Misc
  [ u"ı", "{\\i}" ],
  [ u"µ", "$\\mu$" ],
  [ u"°", "$\\deg$" ],
  [ u"‘", "`" ],    #Quotes
  [ u"’", "'" ],
  [ u"“", "``" ],
  [ u"”", "''" ],
  [ u"‚", "," ],
  [ u"„", ",," ],
]
def replace_accents(string):
    for accent in latexAccents[:]:
        # print(accent)
        string = string.replace(accent[0], accent[1])
    return string