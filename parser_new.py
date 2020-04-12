import pandas as pd
import numpy as np
import column_names_mapping as mapper
import math

from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import date
from timeit import default_timer as timer


def excel_write(
    df, sheet_name, startrow, columns, date_format, datetime_format
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
        columns=columns,
        startrow=startrow,
        header=False,
        index=False,
    )
    writer.save()


def action_check(df, value, output_col_name):
    actions = {
        '&': merge,
        'dict_replace': dict_replace,
        'm': move,
        'ssn': ssn,
        'w': write,
        'collection_status': collection_status,
        'aging_bucket': aging_bucket,
        'client_name': client_name,
        'acc_num': acc_num,
    }
    return actions[value.action](df, value, output_col_name)


def merge(df, value, output_col_name):
    df[output_col_name] = ''
    df[output_col_name] = df[output_col_name].str.cat(
        df[value.old_position].astype(str), sep=value.sep,
    )
    return df[output_col_name]


def dict_replace(df, value, output_col_name):
    df[output_col_name] = df[value.old_position].replace(value.add_info)
    return df[output_col_name]


def move(df, value, output_col_name):
    df[output_col_name] = df[value.old_position]
    return df[output_col_name]


def ssn(df, value, output_col_name):
    base_col_len = len(df[value.old_position[0]].iloc[0].astype(str))
    add_info_len = len(value.add_info)
    n = '0' * (9 - base_col_len - add_info_len)
    df['ssn'] = value.add_info
    df['0s'] = n
    df[output_col_name] = ''
    df[output_col_name] = df[value.old_position[0]].astype(str).str.cat(
        [df['0s'], df['ssn']]
    )
    return df[output_col_name]


def write(df, value, output_col_name):
    df[output_col_name] = pd.Series(
        data=value.add_info, 
        index=range(mapper.settings['total_rows']), 
        name=value.new_position,
    )
    return df[output_col_name]


def collection_status(df, value, output_col_name):
    ocn = output_col_name
    col_idx = value.old_position[0]
    now = pd.to_datetime('now')
    df[ocn] = (now - df[col_idx]).dt.total_seconds() / (60*60*24*365.25)
    masks = [(df[ocn] >= 0) & (df[ocn] <= 30),
             (df[ocn] > 30) & (df[ocn] <= 60),
             (df[ocn] > 60)]
    vals = ['STATEMENT 1', 'STATEMENT 2', 'LETTER 26']
    df[ocn] = np.select(masks, vals, default=0)
    return df[ocn]


def aging_bucket(df, value, output_col_name):
    ocn = output_col_name
    col_idx = value.old_position[0]
    now = pd.to_datetime('now')
    df[ocn] = (now - df[col_idx]).dt.total_seconds() / (60*60*24)
    masks = [
        (df[ocn] >= 0) & (df[ocn] <= 30),
        (df[ocn] > 30) & (df[ocn] <= 60),
        (df[ocn] > 60) & (df[ocn] <= 90),
        (df[ocn] > 90) & (df[ocn] <= 120),
        (df[ocn] > 120) & (df[ocn] <= 360),
        (df[ocn] > 360) & (df[ocn] <= 9999)
    ]
    vals = ['0-30', '31-60', '61-90', '91-120', '121-360', '361-9999']
    df[ocn] = np.select(masks, vals, default=0)
    return df[ocn]


def client_name(df, value, output_col_name):
    ocn = output_col_name
    col_idx = value.old_position[0]
    now = pd.to_datetime('now')
    df[ocn] = (now - df[col_idx]).dt.total_seconds() / (60*60*24)
    masks = [
        (df[ocn] >= 0) & (df[ocn] <= 30),
        (df[ocn] > 30) & (df[ocn] <= 60),
        (df[ocn] > 60) & (df[ocn] <= 90),
        (df[ocn] > 90) & (df[ocn] <= 120),
        (df[ocn] > 120) & (df[ocn] <= 210),
        (df[ocn] > 210) & (df[ocn] <= 300),
        (df[ocn] > 300) & (df[ocn] <= 360),
        (df[ocn] > 360) & (df[ocn] <= 9999)
    ]
    vals = [
        '010E01', '010E02', '010PR1', '010LT1', '010LT2', '010LT2A',
        '010LT2B', '010LT3'
    ]
    df[ocn] = np.select(masks, vals, default=0)
    return df[ocn]


def acc_num(key, value):
    ocn = output_col_name
    col0_idx = value.old_position[0]
    col1_idx = value.old_position[1]
    df[col1_idx] = df[col1_idx].dt.strftime('%m.%d.%Y')
    df = df.astype(str)
    df['eml'] = 'EML'
    df[col0_idx] = df[col0_idx].str.cat(
        df[[col1_idx, 'eml']], sep=value.sep,
    )
    return df[col0_idx]


def main():     
    df = pd.read_excel(
        io=mapper.file_paths['input_file'],
        header=None,
        skiprows=mapper.settings['first_rows_skipped'], 
        keep_default_na=False,
    )
    for key, value in mapper.relationships.items():
        action_check(df, value, key)
    excel_write(
        df=df,
        sheet_name='Parsed data',
        columns=list(mapper.relationships.keys()),
        startrow=mapper.settings['first_writing_row'],
        date_format=value.date_format,
        datetime_format=value.datetime_format,
    )

if __name__ == "__main__":
    main()


