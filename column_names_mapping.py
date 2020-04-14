from dataclasses import make_dataclass
from openpyxl import Workbook
from pathlib import Path
from pprint import pprint


import pandas as pd

# README
# To specify a colum position in input Excel file just write
# 'Column Title': 'Colum Name'
# for eg.
# 'icd10claimdiagdescr01': 'R'
#
# relationships:
# place in list [
# 'columns which compose on the key', 
# 'action identicator',
# 'option: separator'
# ]
actions = {
    '&': 'merge',
    'dict_replace': 'dict_replace',
    'm': 'move',
    'w': 'write',
    'ssn': 'ssn',
}


relationships = {
    # 'icd10claimdiagdescr01': {
    #     'columns': 'R/S', 'action': '&', 'sep': '',
    # },
    # 'icd10claimdiagdescr02': {
    #     'columns': 'T/U', 'action': '&', 'sep': '',
    # },
    # 'icd10claimdiagdescr03': {
    #     'columns': 'V/W', 'action': '&', 'sep': '',
    # },
    # 'svc dept bill name': {
    #     'columns': 'C', 'action': 'dict_replace', 
    #     'add_info': {
    #         'Oklahoma City': 'Echelon Medical',
    #         'Oklahoma City BP': 'The Brace Place',
    #         'Oklahoma City FS': 'First Steps Orthotics',
    #         'Tulsa': 'Echelon Medical',
    #         'Tulsa BP': 'The Brace Place',
    #         'Tulsa FS': 'First Steps Orthotics',
    #         'Medical Motion': 'Medical Motion',
    #     },
    # },
    # 'patient address': {
    #     'columns': 'H/I/J/K/L', 'action': '&',
    # },
    # 'patient address1': {
    #     'columns': 'H', 'action': 'm',
    # },
    # 'patient address2': {
    #     'columns': 'I', 'action': 'm',
    # },
    # 'patient city':{
    #     'columns': 'F', 'action': 'm',
    # },
    # 'patient state': {
    #     'columns': 'K', 'action': 'm',
    # },
    # 'patient zip': {
    #     'columns': 'L', 'action': 'm',
    # },
    # 'patientdob': {
    #     'columns': 'P', 'action': 'm', 'datetime_format': 'DD-MM-YYYY'
    # },
    # 'patient firstname': {
    #     'columns': 'F', 'action': 'm',
    # },
    # 'patient lastname': {
    #     'columns': 'E', 'action': 'm',
    # },
    # 'guarantor addr': {
    #     'columns': 'AT', 'action': 'm',
    # },
    # 'guarantor addr2': {
    #     'columns': 'AU', 'action': 'm',
    # },
    # 'guarantor city': {
    #     'columns': 'AV', 'action': 'm',
    # },
    # 'guarantor email': {
    #     'columns': 'BA', 'action': 'm',
    # },
    # 'guarantor frstnm': {
    #     'columns': 'AS', 'action': 'm',
    # },
    # 'guarantor lastnm': {
    #     'columns': 'AR', 'action': 'm',
    # },
    # 'guarantor phone': {
    #     'columns': 'AY', 'action': 'm',
    # },
    # 'ptnt grntr rltnshp': {
    #     'columns': 'AQ', 'action': 'm',
    # },
    # 'guarantor state': {
    #     'columns': 'AW', 'action': 'm',
    # },
    # 'guarantor zip': {
    #     'columns': 'AX', 'action': 'm',
    # },
    # 'patient homephone': {
    #     'columns': 'M', 'action': 'm',
    # },
    # 'patientid': {
    #     'columns': 'O', 'action': 'm',
    # },
    # 'patient middleinitial': {
    #     'columns': 'G', 'action': 'm',
    # },
    # 'patient mobile no': {
    #     'columns': 'N', 'action': 'm',
    # },
    # 'proccode-descr': {
    #     'columns': 'AJ/AB', 'action': '&', 'sep': '-',
    # },
    # 'guarantor ssn': {
    #     'columns': 'O', 'action': 'ssn', 'add_info': '110',
    # },
    # 'patient ssn': {
    #     'columns': 'O', 'action': 'ssn', 'add_info': '010',
    # },
    # 'Ordering Physician': {
    #     'columns': 'AM', 'action': 'm',
    # },
    # 'invid': {
    #     'columns': 'A', 'action': 'm',
    # },
    # 'postdate': {
    #     'columns': 'B', 'action': 'm',
    # },
    # 'srvdate': {
    #     'columns': 'AC', 'action': 'm',
    # },
    # 'Discount Threshold': {
    #     'columns': '', 'action': 'w', 'add_info': '30% W/O Mgt Approval',
    # },
    # 'Client Billing System': {
    #     'columns': '', 'action': 'w', 'add_info': 'Brightree',
    # },
    # 'Collection Status': {
    #     'columns': 'P', 'action': 'collection_status',
    # },
    # 'Client Billing System User/Pass': {
    #     'columns': '', 'action': 'w', 'add_info': 'See Management',
    # },
    # 'Accepted Payment Forms': {
    #     'columns': '', 'action': 'w', 
    #     'add_info': 'Credit, Debit, e-Check, Mail-In',
    # },
    # 'Financial Class': {
    #     'columns': '', 'action': 'w', 'add_info': 'Patient Responsibility',
    # },
    # 'Client Billing System URL': {
    #     'columns': '', 'action': 'w', 
    #     'add_info': 'https://login.brightree.net/',
    # },
    # 'Responsibility Date': {
    #     'columns': 'B', 'action': 'm',
    # },
    # 'Client Name': {
    #     'columns': 'B', 'action': 'client_name',
    # },
    # '3rd Party Correspondence': {
    #     'columns': '', 'action': 'w', 
    #     'add_info': 'Innovare-Virtual Post Mail',
    # },
    # 'Script': {
    #     'columns': 'Client Name', 'action': 'dict_replace', 
    #     'add_info': {
    #         '010E01': 'Statement',
    #         '010E02': 'Statement',
    #         '010PR1': 'Account Review 1',
    #         '010LT1': 'Account Review 2',
    #         '010LT2': 'Account Review 3',
    #         '010LT2A': 'Settlement Offer',
    #         '010LT2B': 'Settlement Offer',
    #         '010LT3': 'Settlement Offer',
    #     },
    # },
    # 'Early Out Correspondence': {
    #     'columns': '', 'action': 'w', 'add_info': 'Managed By Client',
    # },
    # 'Client Payment Mailing Address': {
    #     'columns': '', 'action': 'w', 
    #     'add_info': '340 S Lemon Ave #1102 Walnut, CA 91789',
    # },
    # 'Client Payment System': {
    #     'columns': '', 'action': 'w', 'add_info': 'Repay',
    # },
    # 'Callback Number': {
    #     'columns': '', 'action': 'w', 'add_info': '(405) 200-1666',
    # },
    # 'Minimum Payment': {
    #     'columns': '', 'action': 'w', 'add_info': '$10,00',
    # },
    # 'Internal Account Status': {
    #     'columns': 'Script', 'action': 'm',
    # },
    # 'Client Phone': {
    #     'columns': '', 'action': 'w', 'add_info': '(405) 200-1666',
    # },
    # 'Billing Provider': {
    #     'columns': 'D', 'action': 'm',
    # },
    # 'Claim Received Date': {
    #     'columns': '', 'action': 'w', 'add_info': pd.to_datetime('now'),
    # },
    # 'Client Billing Contact': {
    #     'columns': '', 'action': 'w', 
    #     'add_info': 'Sinthya Cruz-Billing Manager',
    # },
    # 'Client Payment System URL': {
    #     'columns': '', 'action': 'w', 
    #     'add_info': 'https://innovareprm.repay.io',
    # },
    # 'Client Website': {
    #     'columns': '', 'action': 'w', 'add_info': 'N/A',
    # },
    # 'Aging Bucket': {
    #     'columns': 'B', 'action': 'aging_bucket', 'add_info': '',
    # },
    # 'Customer Service Email': {
    #     'columns': '', 'action': 'w', 'add_info': 'support@innovareprm.com',
    # },
    # 'Specialty': {
    #     'columns': '', 'action': 'w', 
    #     'add_info': 'Sleep Medicine and Supplies',
    # },
    # 'Custom Account Number': {
    #     'columns': 'O/AC', 'action': 'acc_num', 'sep': '-'
    # },
    # 'charge off date': {
    #     'columns': 'B', 'action': 'm',
    # },
    # 'originated date': {
    #     'columns': 'AC', 'action': 'm',
    # },
    # 'Patient Address': {
    #     'columns': 'H/I/J/K/L', 'action': '&',
    # },
    # 'Phone Number1': {
    #     'columns': 'M', 'action': 'm',
    # },
    # 'Phone Number2': {
    #     'columns': 'N', 'action': 'm',
    # },
    # 'Phone Number3': {
    #     'columns': 'AY', 'action': 'm',
    # },
    # 'Phone Number4': {
    #     'columns': 'AZ', 'action': 'm',
    # },
    # 'creditor': {
    #     'columns': '', 'action': 'w', 
    #     'add_info': 'Echelon Medical',
    # },
    # 'Action Code': {
    #     'columns': 'AG', 'action': 'action_code', 
    # },
    # 'original claim amount (DOS Rows)': '',
    # 'Balance (DOS Rows)': '',
    # r'10% discount': '',
    # r'15% discount': '',
    # r'20% discount': '',
    # r'25% discount': '',
    # r'30% discount': '',
    # 'original claim amount (Totals Row)': '',
    # 'Balance (Totals Row)': '',
    'Invoice Detail Charge': {
        'columns': 'AE', 'action': 'currency',
    },
    'Invoice Detail Allow': {
        'columns': 'AF', 'action': 'currency',
    },
    'Invoice Detail Payments': {
        'columns': 'AG', 'action': 'currency',
    },
    'Invoice Detail Adjustments': {
        'columns': 'AH', 'action': 'currency',
    },
    'Invoice Detail Balance': {
        'columns': 'AI', 'action': 'currency',
    },
}


settings = {
    'input_file_name': 'data_to_parse',
    'output_file_name': 'Parsed data',
    # rows to skip in input file for eg. headeers
    # integers need to be placed in list, 0 is always the first row/column
    'first_rows_skipped': [0, 1, 2],
    'cols_to_read': 'A:BA',
    # row from which u want to start filling excel
    # for eg. u want to skip headers so we start writing from first row
    'first_writing_row': 1,
    'total_rows': 6336,
}


file_paths = {
    'input_file': (
        Path.cwd() / 'file_to_process' / f'{settings["input_file_name"]}.xlsx'
    ),
    'output_file': (
        Path.cwd() / 'processed_file' / f'{settings["output_file_name"]}.xlsx'
    ),
}


DataKey = make_dataclass(
    'DataKey',
    ['old_position', 'new_position', 'action', 'sep', 'add_info', 
    'date_format', 'datetime_format']
)


def col2num(col):
    num = 0 
    for c in col:
        num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num - 1


new_pos = 0
for key, value in relationships.items():
    column_indexes = []
    action = value.get('action', '') 
    sep = value.get('sep', ' ') 
    add_info = value.get('add_info', '')
    date_format = value.get('date_format', None)
    datetime_format = value.get('datetime_format', None)
    for char in value['columns'].split('/'):
        if len(char) < 3:
            column_indexes.append(col2num(char))
        else:
            column_indexes.append(char)
    relationships[key] = DataKey(
        old_position=column_indexes,
        new_position=new_pos,
        action=action,
        sep=sep,
        add_info=add_info,
        date_format=date_format,
        datetime_format=datetime_format,
    )
    new_pos += 1


wb = Workbook()
ws = wb.active
ws.title = settings['output_file_name']
ws.append(list(relationships.keys()))
for column_cells in ws.columns:
    length = max(len(cell.value) + 5 for cell in column_cells)
    ws.column_dimensions[column_cells[0].column_letter].width = length
wb.save(filename = file_paths['output_file'])

pprint(relationships)

# alphabet_to_num = {
#     'A': 0,
#     'B': 1,
#     'C': 2,
#     'D': 3,
#     'E': 4,
#     'F': 5,
#     'G': 6,
#     'H': 7,
#     'I': 8,
#     'J': 9,
#     'K': 10,
#     'L': 11,
#     'M': 12,
#     'N': 13,
#     'O': 14,
#     'P': 15,
#     'Q': 16,
#     'R': 17,
#     'S': 18,
#     'T': 19,
#     'U': 20,
#     'V': 21,
#     'W': 22,
#     'X': 23,
#     'Y': 24,
#     'Z': 25,
#     'AA': 26,
#     'AB': 27,
#     'AC': 28,
#     'AD': 29,
#     'AE': 30,
#     'AF': 31,
#     'AG': 32,
#     'AH': 33,
#     'AI': 34,
#     'AJ': 35,
#     'AK': 36,
#     'AL': 37,
#     'AM': 38,
#     'AN': 39,
#     'AO': 40,
#     'AP': 41,
#     'AQ': 42,
#     'AR': 43,
#     'AS': 44,
#     'AT': 45,
#     'AU': 46,
#     'AV': 47,
#     'AW': 48,
#     'AX': 49,
#     'AY': 50,
#     'AZ': 51,
#     'BA': 52,
#     'BB': 53,
#     'BC': 54,
#     'BD': 55,
#     'BE': 56,
#     'BF': 57,
#     'BG': 58,
#     'BH': 59,
#     'BI': 60,
#     'BJ': 61,
#     'BK': 62,
#     'BL': 63,
#     'BM': 64,
#     'BN': 65,
#     'BO': 66,
#     'BP': 67,
#     'BQ': 68,
#     'BR': 69,
#     'BS': 70,
#     'BT': 71,
#     'BU': 72,
#     'BV': 73,
#     'BW': 74,
#     'BX': 75,
#     'BY': 76,
#     'BZ': 77,
#     'CA': 78,
#     'CB': 79,
#     'CC': 80,
#     'CD': 81,
#     'CE': 82,
#     'CF': 83,
# }