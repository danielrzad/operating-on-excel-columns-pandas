from dataclasses import make_dataclass
from openpyxl import Workbook
from pathlib import Path
from pprint import pprint


import parser


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
    's': 'svc',
    'm': 'move'
}


relationships = {
    'icd10claimdiagdescr01': ['R/S', '&', ' '],
    'icd10claimdiagdescr02': ['T/U', '&', ''],
    'icd10claimdiagdescr03': ['V/W', '&', ''],
    'svc dept bill name': ['C', 's', ''],
    'patient address': ['H/I/J/K/L', '&', ' '],
    'patient address1': ['H', 'm', ''],
    'patient address2': ['I', 'm', ''],
    'patient city': ['J', 'm', ''],
    'patient state': ['K', 'm', ''],
    'patient zip': ['L', 'm', ''],
    'patientdob': ['P', 'm', ''],
    'patient firstname': ['F', 'm', ''],
    'patient lastname': ['E', 'm', ''],
    'guarantor addr': ['AT', 'm', ''],
    'guarantor addr2': ['AU', 'm', ''],
    'guarantor city': ['AV', 'm', ''],
    'guarantor email': ['BA', 'm', ''],
    'guarantor frstnm': ['AS', 'm', ''],
    'guarantor lastnm': ['AR', 'm', ''],
    'guarantor phone': ['AY', 'm', ''],
    'ptnt grntr rltnshp': ['AQ', 'm', ''],
    'guarantor state': ['AW', 'm', ''],
    'guarantor zip': ['AX', 'm', ''],
    'patient homephone': ['M', 'm', ''],
    'patientid': ['O', 'm', ''],
    'patient middleinitial': ['G', 'm', ''],
    'patient mobile no': ['N', 'm', ''],
    'proccode-descr': ['AJ/AB', '&', '-'],
    # 'guarantor ssn': ['M', 'm', ''],
    # 'patient ssn': ['M', 'm', ''],
    # 'Ordering Physician': ['M', 'm', ''],
    # 'invid': ['M', 'm', ''],
    # 'postdate': ['M', 'm', ''],
    # 'srvdate': ['M', 'm', ''],
    # 'Discount Threshold': ['M', 'm', ''],
    # 'Client Billing System': ['M', 'm', ''],
    # 'Collection Status': ['M', 'm', ''],
    # 'Client Billing System User/Pass': ['M', 'm', ''],
    # # 'Accepted Payment Forms': '',
    # 'Financial Class': '',
    # 'Client Billing System URL': '',
    # 'Responsibility Date': '',
    # 'Client Name': '',
    # '3rd Party Correspondence': '',
    # 'Script': '',
    # 'Early Out Correspondence': '',
    # 'Client Payment Mailing Address': '',
    # 'Client Payment System': '',
    # 'Callback Number': '',
    # 'Minimum Payment': '',
    # 'Internal Account Status': '',
    # 'Client Phone': '',
    # 'Billing Provider': '',
    # 'Claim Received Date': '',
    # 'Client Billing Contact': '',
    # 'Client Payment System URL': '',
    # 'Client Website': '',
    # 'Aging Bucket': '',
    # 'Customer Service Email': '',
    # 'Specialty': '',
    # 'Custom Account Number': '',
    # 'charge off date': '',
    # 'originated date': '',
    # 'Patient Address': '',
    # 'Phone Number1': '',
    # 'Phone Number2': '',
    # 'Phone Number3': '',
    # 'Phone Number4': '',
    # 'creditor': '',
    # 'Action Code': '',
    # 'original claim amount (DOS Rows)': '',
    # 'Balance (DOS Rows)': '',
    # r'10% discount': '',
    # r'15% discount': '',
    # r'20% discount': '',
    # r'25% discount': '',
    # r'30% discount': '',
    # 'original claim amount (Totals Row)': '',
    # 'Balance (Totals Row)': '',
    # 'Invoice Detail Charge': '',
    # 'Invoice Detail Allow': '',
    # 'Invoice Detail Payments': '',
    # 'Invoice Detail Adjustments': '',
    # 'Invoice Detail Balance': '',
}


settings = {
    'input_file_name': 'data_to_parse',
    'output_file_name': 'Parsed data',
    # rows to skip in input file for eg. headeers
    # integers need to be placed in list, 0 is always the first row/column
    'first_rows_skipped': [0, 1, 2], 
    # row from which u want to start filling excel
    # for eg. u want to skip headers so we start writing from first row
    'first_writing_row': 1, 
}


file_paths = {
    'input_file': (
        Path.cwd() / 'file_to_process' / f'{settings["input_file_name"]}.xlsx'
    ),
    'output_file': (
        Path.cwd() / 'processed_file' / f'{settings["output_file_name"]}.xlsx'
    ),
}


alphabet_to_num = {
    'A': 0,
    'B': 1,
    'C': 2,
    'D': 3,
    'E': 4,
    'F': 5,
    'G': 6,
    'H': 7,
    'I': 8,
    'J': 9,
    'K': 10,
    'L': 11,
    'M': 12,
    'N': 13,
    'O': 14,
    'P': 15,
    'Q': 16,
    'R': 17,
    'S': 18,
    'T': 19,
    'U': 20,
    'V': 21,
    'W': 22,
    'X': 23,
    'Y': 24,
    'Z': 25,
    'AA': 26,
    'AB': 27,
    'AC': 28,
    'AD': 29,
    'AE': 30,
    'AF': 31,
    'AG': 32,
    'AH': 33,
    'AI': 34,
    'AJ': 35,
    'AK': 36,
    'AL': 37,
    'AM': 38,
    'AN': 39,
    'AO': 40,
    'AP': 41,
    'AQ': 42,
    'AR': 43,
    'AS': 44,
    'AT': 45,
    'AU': 46,
    'AV': 47,
    'AW': 48,
    'AX': 49,
    'AY': 50,
    'AZ': 51,
    'BA': 52,
    'BB': 53,
    'BC': 54,
    'BD': 55,
    'BE': 56,
    'BF': 57,
    'BG': 58,
    'BH': 59,
    'BI': 60,
    'BJ': 61,
    'BK': 62,
    'BL': 63,
    'BM': 64,
    'BN': 65,
    'BO': 66,
    'BP': 67,
    'BQ': 68,
    'BR': 69,
    'BS': 70,
    'BT': 71,
    'BU': 72,
    'BV': 73,
    'BW': 74,
    'BX': 75,
    'BY': 76,
    'BZ': 77,
    'CA': 78,
    'CB': 79,
    'CC': 80,
    'CD': 81,
    'CE': 82,
    'CF': 83,
}


DataKey = make_dataclass(
    'DataKey', ['old_position', 'new_position', 'action', 'sep']
)

new_pos = 0
for key, value in relationships.items():
    if isinstance(value, int):
        continue
    column_indexes = []
    action = ''
    sep = value[2]
    for char in value[0].split('/'):
        if char in alphabet_to_num:
            print(value[0].split())
            print(char)
            column_indexes.append(alphabet_to_num[char])
    if value[1] in actions:
        action = actions[value[1]]
    print(column_indexes)
    relationships[key] = DataKey(
        old_position=column_indexes,
        new_position=new_pos,
        action=action,
        sep=sep,
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