import json
import pandas
import re
from pathlib import Path

pandas.set_option('display.max_columns', None)

# file path handling
HERE = Path(__file__).parent.resolve()
INPUT = HERE.parent.joinpath('data', 'CTDB_form_legend.xlsx')
OUTPUT_DIR = HERE.parent.joinpath('phenotype')
status = OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# dictionary of dataframes
try:
    legends = pandas.read_excel(INPUT, sheet_name=None)
except PermissionError as err:
    print(f"Close your spreadsheet, {INPUT}, and try again.")
    raise err

columns = []
for legend in legends.keys():
    df = legends[legend]
    columns += list(df.columns)

header = list(set(columns))

for legend in legends.keys():
    df = legends[legend]
    df.drop(df.shape[0]-1)

    formname = re.sub( r'[\W]+' , '' , legend.replace(' ', '_') ).lower()

    out_filename = OUTPUT_DIR.joinpath(f'{formname}.json')

    qid_uniq = df['QUESTIONID'].unique()
    qids = qid_uniq[qid_uniq != ' '].astype('int')

    d = {}
    current = ''
    for i, row in df.iterrows():
        # detecting if on first line of data dictionary/legend
        if current == '':
            # start
            qid = str(int(row['QUESTIONID']))
            current = qid
            if pandas.isna(row['QUESTION_ALIAS']):
                ShortName = row['QUESTION_NAME']
            else:
                ShortName = row['QUESTION_ALIAS']
            LongName = row['QUESTION_NAME'] + ' (question ID ' + qid + ')'
            Description = row['QUESTION_TEXT']
            # DataType = row[10]
            # Range = row[12]
            Levels = None

        elif not pandas.isna(row['QUESTION_TEXT']):
            # write
            d[ShortName] = {}
            d[ShortName]['LongName'] = LongName
            d[ShortName]['Description'] = Description
            if Levels:
                d[ShortName]['Levels'] = Levels
            # if DataType == 'Numeric':
            #     d[ShortName]['Range'] = Range

            # reset
            qid = str(int(row['QUESTIONID']))
            if pandas.isna(row['QUESTION_ALIAS']):
                ShortName = row['QUESTION_NAME']
            else:
                ShortName = row['QUESTION_ALIAS']
            LongName = row['QUESTION_NAME'] + ' (question ID ' + qid + ')'
            Description = row['QUESTION_TEXT']
            # DataType = row[10]
            # Range = row[12]
            Levels = None

        if not Levels:
            if not pandas.isna(row['CODEVALUE']):
                if type(row['DISPLAY']) == float:
                    if pandas.isna(row['DISPLAY']):
                        Levels = {str(int(row['CODEVALUE'])): 'N/A'}
                    else:
                        Levels = {str(int(row['CODEVALUE'])): str(int(row['DISPLAY']))}
                else:
                    Levels = {str(int(row['CODEVALUE'])): str(row['DISPLAY'])}
        else:
            if not pandas.isna(row['CODEVALUE']):
                if type(row['DISPLAY']) == float:
                    if pandas.isna(row['DISPLAY']):
                        Levels[str(int(row['CODEVALUE']))] = 'N/A'
                    else:
                        Levels[str(int(row['CODEVALUE']))] = str(int(row['DISPLAY']))
                else:
                    Levels[str(int(row['CODEVALUE']))] = str(row['DISPLAY'])

        # detecting if on last line of data dictionary/legend
        if i == df.shape[0] - 1:
            # write
            d[ShortName] = {}
            d[ShortName]['LongName'] = LongName
            d[ShortName]['Description'] = Description
            if Levels:
                d[ShortName]['Levels'] = Levels
            # if DataType == 'Numeric':
            #     d[ShortName]['Range'] = Range

    with open(out_filename, 'w') as f:
        f.write(json.dumps(d, indent=4))