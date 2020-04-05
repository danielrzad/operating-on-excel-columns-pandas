import pandas as pd
import numpy as np
import column_names_mapping as mapper


from openpyxl import Workbook
from openpyxl import load_workbook
from timeit import default_timer as timer


# co do zrobienia
# zajac sie Collection Status

def excel_write(
    df, sheet_name, startrow, startcol, date_format, datetime_format
):
    file = mapper.file_paths['output_file']
    wb = load_workbook(file)
    writer = pd.ExcelWriter(
        file, mode='a', date_format=date_format, 
        datetime_format=datetime_format,
    )
    writer.book = wb
    writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)
    df.to_excel(
        excel_writer=writer,
        sheet_name=sheet_name,
        startrow=startrow,
        startcol=startcol,
        header=False,
        index=False,
    )
    writer.save()


def excel_read(
    io, header, names, skiprows, usecols, keep_default_na,
    ):
    df = pd.read_excel(
            io=io,
            header=header,
            names=names,
            skiprows=skiprows,
            usecols=usecols,
            keep_default_na=keep_default_na,
    )
    return df



def action_check(value):
    actions = {
    'merge': merge,
    'svc': svc,
    'm': move,
    'ssn': ssn,
    'w': write,
    }
    return actions[value.action](value)


def merge(value):
    df = excel_read(
        io=mapper.file_paths['input_file'], 
        header=None, 
        names=value.old_position,
        skiprows=mapper.settings['first_rows_skipped'], 
        usecols=value.old_position, 
        keep_default_na=False,
    )
    df['merged'] = ''
    print(df)
    df = df.astype(str)
    df['merged'] = df['merged'].str.cat(
        df[value.old_position], sep=value.sep,
    )
    print(df['merged'])
    print(df.dtypes)
    return df['merged']


def svc(value):
    df = excel_read(
        io=mapper.file_paths['input_file'], 
        header=None, 
        names=value.old_position,
        skiprows=mapper.settings['first_rows_skipped'], 
        usecols=value.old_position, 
        keep_default_na=False,
    )
    cities = {
    'Oklahoma City': 'Echelon Medical',
    'Oklahoma City BP': 'The Brace Place',
    'Oklahoma City FS': 'First Steps Orthotics',
    'Tulsa': 'Echelon Medical',
    'Tulsa BP': 'The Brace Place',
    'Tulsa FS': 'First Steps Orthotics',
    'Medical Motion': 'Medical Motion',
    }
    df = df.replace(cities)
    print(df)
    print(df.dtypes)
    return df


def move(value):
    df = excel_read(
        io=mapper.file_paths['input_file'], 
        header=None, 
        names=value.old_position,
        skiprows=mapper.settings['first_rows_skipped'], 
        usecols=value.old_position, 
        keep_default_na=False,
    )
    print(df)
    print(df.dtypes)
    return df


def ssn(value):
    df = excel_read(
        io=mapper.file_paths['input_file'], 
        header=None, 
        names=value.old_position,
        skiprows=mapper.settings['first_rows_skipped'], 
        usecols=value.old_position, 
        keep_default_na=False,
    )
    base_col_len = len(df[value.old_position[0]].iloc[0].astype(str))
    add_info_len = len(value.add_info)
    n = '0' * (9 - base_col_len - add_info_len)
    df['ssn'] = value.add_info
    df['0s'] = n
    df[value.old_position[0]] = df[value.old_position[0]].astype(str).str.cat(
        [df['0s'], df['ssn']])
    print(df[value.old_position[0]])
    print(df.dtypes)
    return df[value.old_position[0]]


def write(value):
    df = pd.Series(
        data=value.add_info, 
        index=range(mapper.settings['total_rows']), 
        name=value.new_position,
    )
    print(df)
    print(df.dtypes)
    return df


for key, value in mapper.relationships.items():
    start = timer()
    excel_write(
        df=action_check(value),
        sheet_name='Parsed data', 
        startrow=mapper.settings['first_writing_row'], 
        startcol=value.new_position,
        date_format=value.date_format, 
        datetime_format=value.datetime_format,
    )
    end = timer()
    print(end - start)

print('DONE')