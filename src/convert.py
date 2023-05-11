import pandas
import re
from pathlib import Path

# file path handling
HERE = Path(__file__).parent.resolve()
INPUT = HERE.parent.joinpath('data', 'HV_CTDB-data-download-20211202.xlsx')
# INPUT = HERE.parent.joinpath('data', 'NNT CTDB Data Download (modified).xlsx')
OUTPUT_DIR = HERE.parent.joinpath('phenotype')
status = OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ignoreable spreadsheets
ignore_sheet_list = [
    "Big-Five Personality Survey",
    "Brief Trauma Questionnaire (BQ)",
    "Consent V2",
    "Email Reminder"
]

try:
    ctdb_dict_of_dfs = pandas.read_excel(INPUT, sheet_name=None)
except PermissionError as err:
    print(f"Close your spreadsheet and try again.")
    raise err

for form, df in ctdb_dict_of_dfs.items():

    if form not in ignore_sheet_list:
        ignore_column_list = []

        for column in df.columns:
            # if the column is subject/visit PII
            if column == ' ' or column.startswith(' .') or column.startswith('Unnamed: '):
                ignore_column_list.append(column)

    cleaned_df = df.drop(ignore_column_list, axis=1).drop(0, axis=0)

    messy_df = df[ignore_column_list]
    messy_df.columns = messy_df.iloc[0]
    messy_df.drop(0, axis=0, inplace=True)

    # convert words and strings to exclusively alphanumerics and underscores for BIDS' sake
    formname = re.sub( r'[\W]+' , '' , form.replace(' ', '_') ).lower()

    cleaned_df.to_csv(OUTPUT_DIR.joinpath(f'{formname}.tsv'), sep='\t', index=False)
    messy_df.to_csv(OUTPUT_DIR.joinpath(f'{formname}_PII.tsv'), sep='\t', index=False)
