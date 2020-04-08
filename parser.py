import pandas as pd
import numpy as np
import column_names_mapping as mapper
import math

from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import date
from timeit import default_timer as timer

# to do list


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
    'collection_status': collection_status,
    'aging_bucket': aging_bucket,
    'client_name': client_name,
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


def collection_status(value):
    df = excel_read(
        io=mapper.file_paths['input_file'], 
        header=None, 
        names=value.old_position,
        skiprows=mapper.settings['first_rows_skipped'], 
        usecols=value.old_position, 
        keep_default_na=False,
    )  
    col_idx = value.old_position[0]
    now = pd.to_datetime('now')
    df[col_idx] = (now - df[col_idx]).dt.total_seconds() // (60*60*24*365.25)
    df.loc[
        (col_idx>=0) & (col_idx<=30), value.old_position[0]
    ] = 'STATEMENT 1'
    df.loc[
        (col_idx>=31) & (col_idx<=60), value.old_position[0]
    ] = 'STATEMENT 2'
    df.loc[(col_idx>61), value.old_position[0]] = "LETTER 26"
    return df


def aging_bucket(value):
    df = excel_read(
        io=mapper.file_paths['input_file'], 
        header=None, 
        names=value.old_position,
        skiprows=mapper.settings['first_rows_skipped'], 
        usecols=value.old_position, 
        keep_default_na=False,
    )
    col_idx = value.old_position[0]
    now = pd.to_datetime('now')
    df[col_idx] = (now - df[col_idx]).dt.total_seconds() // (60*60*24)
    df.loc[
        (df[col_idx]>=0) & (df[col_idx]<=30), value.old_position[0]
    ] = '0-30'
    df.loc[
        (df[col_idx]>=31) & (df[col_idx]<=60), value.old_position[0]
    ] = '31-60'
    df.loc[
        (df[col_idx]>=61) & (df[col_idx]<=90), value.old_position[0]
    ] = '61-90'
    df.loc[
        (df[col_idx]>=91) & (df[col_idx]<=120), value.old_position[0]
    ] = '91-120'
    df.loc[
        (df[col_idx]>=121) & (df[col_idx]<=360), value.old_position[0]
    ] = '121-360'
    df.loc[
        (df[col_idx]>=361) & (df[col_idx]<=9999), value.old_position[0]
    ] = '361-9999'
    return df


def client_name(value):
    df = excel_read(
        io=mapper.file_paths['input_file'], 
        header=None, 
        names=value.old_position,
        skiprows=mapper.settings['first_rows_skipped'], 
        usecols=value.old_position, 
        keep_default_na=False,
    )
    days_range = {
        '0-30': '010E01',
        '31-60': '010E02',
        '61-90': '010PR1',
        '91-120': '010LT1',
        '121-210': '010LT2',
        '211-300': '010LT2A',
        '301-360': '010LT2B',
        '361-9999': '010LT3',
    }
    col_idx = value.old_position[0]
    now = pd.to_datetime('now')
    df[col_idx] = (now - df[col_idx]).dt.total_seconds() // (60*60*24)
    df.loc[
        (df[col_idx]>=0) & (df[col_idx]<=30), value.old_position[0]
    ] = '0-30'
    df.loc[
        (df[col_idx]>=31) & (df[col_idx]<=60), value.old_position[0]
    ] = '31-60'
    df.loc[
        (df[col_idx]>=61) & (df[col_idx]<=90), value.old_position[0]
    ] = '61-90'
    df.loc[
        (df[col_idx]>=91) & (df[col_idx]<=120), value.old_position[0]
    ] = '91-120'
    df.loc[
        (df[col_idx]>=121) & (df[col_idx]<=360), value.old_position[0]
    ] = '121-360'
    df.loc[
        (df[col_idx]>=211) & (df[col_idx]<=300), value.old_position[0]
    ] = '211-300'
    df.loc[
        (df[col_idx]>=301) & (df[col_idx]<=360), value.old_position[0]
    ] = '301-360'
    df.loc[
        (df[col_idx]>=361) & (df[col_idx]<=9999), value.old_position[0]
    ] = '361-9999'
    df = df.replace(days_range)
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
