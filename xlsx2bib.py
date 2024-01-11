import pandas as pd
from os.path import splitext
import numpy as np
from pandas.api.types import is_any_real_numeric_dtype


def check_page_range(pr):
    if is_any_real_numeric_dtype(type(pr)) and np.isnan(pr):
        return False
    else:
        return True


def xlsx2bib(xlsx_file):
    # Replace file extension
    base_name = splitext(xlsx_file)[0]
    bib_file = base_name + '.bib'

    # Parse excel file
    df = pd.read_excel(xlsx_file).sort_values(by='Number')

    # For each row in the xlsx file, generated an entry in the bib file
    with open(bib_file, 'w') as fh:
        for i, row in df.iterrows():
            fh.write(f"""@article{{ref{row['Number']},
            author={{{row['First author'].replace('&', chr(92)+'&').strip()}}},
            journal={{{row['Journal'].strip()}}},
            number={{{int(row['Issue']) if not np.isnan(row['Issue']) else ''}}},
            pages={{{str(row['page range']).replace('-', '--').strip() if check_page_range(row['page range']) else ''}}},
            year={{{row['year']}}},
}}

""")
