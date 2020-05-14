from datetime import date


import pandas as pd
import numpy as np
import column_names_mapping as mapper


def excel_write(
    df, file, startrow, columns, header, date_format, datetime_format
):
    writer = pd.ExcelWriter(
        file, date_format=date_format, 
        datetime_format=datetime_format,
    )
    df.to_excel(
        excel_writer=writer,
        columns=columns,
        startrow=startrow,
        header=header,
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
        'action_code': action_code,
        'currency': currency,
    }
    return actions[value.action](df, value, output_col_name)


def merge(df, value, output_col_name):
    df[output_col_name] = df[value.old_position[0]].str.cat(
        df[value.old_position[1:]].astype(str), sep=value.sep,
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
    df[output_col_name] = value.add_info
    return df[output_col_name]


def collection_status(df, value, output_col_name):
    ocn = output_col_name
    col_idx = value.old_position[0]
    now = pd.to_datetime('today')
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
    now = pd.to_datetime('today')
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
    now = pd.to_datetime('today')
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


def acc_num(df, value, output_col_name):
    ocn = output_col_name
    col0_idx = value.old_position[0]
    col1_idx = value.old_position[1]
    df[ocn] = value.add_info
    df[col1_idx] = df[col1_idx].dt.strftime('%m.%d.%Y')
    df['eml'] = 'EML'
    df[ocn] = df[col0_idx].astype(str).str.cat(
        df[[col1_idx, 'eml']], sep=value.sep,
    )
    return df[ocn]


def action_code(df, value, output_col_name):
    ocn = output_col_name
    col_idx = value.old_position[0]    
    df[ocn] = df[value.old_position[0]]
    masks = [
        (df[ocn] == 0)
    ]
    vals = ['INFO ACCOUNT']
    df[ocn] = np.select(masks, vals, default='CORRESPONDENCE ACCOUNT')
    return df[output_col_name]


def format_currency(x):
    thousands_separator = " "
    fractional_separator = ","
    x = '${:,.2f}'.format(x) 
    main_currency = x.split('.')[0]
    fractional_currency = x.split('.')[1]
    new_main_currency = main_currency.replace(',', '.')
    x = new_main_currency + fractional_separator + fractional_currency
    return x


def currency(df, value, output_col_name):
    ocn = output_col_name
    col_idx = value.old_position[0]   
    df[ocn] = df[col_idx].apply(format_currency)
    return df[ocn]


def main():
    for file in mapper.settings['input_files']:
        print('Processing', file)
        df = pd.read_excel(
            io=file,
            header=None,
            skiprows=mapper.settings['first_rows_skipped'], 
            keep_default_na=False,
        )
        for key, value in mapper.relationships.items():
            action_check(df, value, key)
            if value.datetime_format != None:
                df[key] = df[key].dt.strftime(value.datetime_format)
        output_file_path = mapper.file_paths['output_file_folder']
        output_file_name = file.name.replace('.xlsx', '_Output.xlsx')
        df.sort_values(
            by=['srvdate','patientid'],
            ascending=True,
            inplace=True,
        )
        df.to_excel(
            excel_writer=output_file_path / output_file_name,
            columns=list(mapper.relationships.keys()),
            index=False,
        )
        print('Finished Processing', file)


if __name__ == "__main__":
    main()
