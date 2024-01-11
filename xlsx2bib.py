import pandas as pd
from os.path import splitext


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
            number={{{row['Issue']}}},
            pages={{{str(row['page range']).replace('-', '--').strip()}}},
            year={{{row['year']}}},
}}

""")
