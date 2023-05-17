import json
import pandas
import re
from pathlib import Path

pandas.set_option('display.max_columns', None)

# file path handling
HERE = Path(__file__).parent.resolve()
# INPUT = HERE.parent.joinpath('data', '.xlsx')
# INPUT = HERE.parent.joinpath('data', 'CTDB_form_legend.xlsx')
# INPUT = HERE.parent.joinpath('data', 'ARI-P_Legend.xlsx')
INPUT = HERE.parent.joinpath('data', 'Demographics_Legend.xlsx')
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

    if 'QUESTIONID' in df.columns:
        questionid = 'QUESTIONID'
        question_alias = 'QUESTION_ALIAS'
        question_name = 'QUESTION_NAME'
        question_text = 'QUESTION_TEXT'
        codevalue = 'CODEVALUE'
        display = 'DISPLAY'
    elif ' Question Id' in df.columns:
        questionid = ' Question Id'
        question_alias = ' Alias'
        question_name = ' Question Name'
        question_text = ' Question Text'
        codevalue = ' Code Value '
        display = ' Answer'

    formname = re.sub( r'[\W]+' , '' , legend.replace(' ', '_') ).lower()

    out_filename = OUTPUT_DIR.joinpath(f'{formname}.json')

    qid_uniq = df[questionid].unique()
    qids = qid_uniq[qid_uniq != ' '].astype('int')

    d = {}
    current = ''
    for i, row in df.iterrows():
        # detecting if on first line of data dictionary/legend
        if current == '':
            # start
            qid = str(int(row[questionid]))
            current = qid
            if pandas.isna(row[question_alias]) or row[question_alias] == ' ':
                ShortName = row[question_name]
            else:
                ShortName = row[question_alias]
            LongName = row[question_name] + ' (question ID ' + qid + ')'
            Description = row[question_text]
            # DataType = row[10]
            # Range = row[12]
            Levels = None

        elif not pandas.isna(row[question_text]) and row[question_text] != ' ':
            # write
            d[ShortName] = {}
            d[ShortName]['LongName'] = LongName
            d[ShortName]['Description'] = Description
            if Levels:
                d[ShortName]['Levels'] = Levels
            # if DataType == 'Numeric':
            #     d[ShortName]['Range'] = Range

            # reset
            qid = str(int(row[questionid]))
            if pandas.isna(row[question_alias]) or row[question_alias] == ' ':
                ShortName = row[question_name]
            else:
                ShortName = row[question_alias]
            LongName = row[question_name] + ' (question ID ' + qid + ')'
            Description = row[question_text]
            # DataType = row[10]
            # Range = row[12]
            Levels = None

        if not Levels:
            if not pandas.isna(row[codevalue]) and row[codevalue] != ' ':
                if type(row[display]) == float:
                    if pandas.isna(row[display]):
                        Levels = {str(int(row[codevalue])): 'N/A'}
                    else:
                        Levels = {str(int(row[codevalue])): str(int(row[display]))}
                else:
                    Levels = {str(int(row[codevalue])): str(row[display])}
        else:
            if not pandas.isna(row[codevalue]):
                if type(row[display]) == float:
                    if pandas.isna(row[display]):
                        Levels[str(int(row[codevalue]))] = 'N/A'
                    else:
                        Levels[str(int(row[codevalue]))] = str(int(row[display]))
                else:
                    Levels[str(int(row[codevalue]))] = str(row[display])

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
