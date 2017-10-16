from openpyxl import load_workbook
import datetime

path_to_source = '..\\dummy source excel.xlsx'
data_sheet_name = 'Sheet1'
constant_dict = {
    'assignment_team_value': 'SUPPLIERNET',
    'closed_status': 'CLOSED',
    'file_name' : 'P10W04 AMR ALL'
}
column_idx_dict = {
    'assignment_team': 6,
    'status': 5,
    'close_period': 18,
    'close_week': 19,
    'close_year': 17,
    'open_year': 9,
    'open_period': 10,
    'open_week': 11,
    'urgency': 3,
    'environment': 4,
    'request_number': 1,
    'committed_date': 31,
    'description': 40
}
source_workbook = load_workbook(filename=path_to_source, read_only=True)
source_worksheet = source_workbook[data_sheet_name]

def check_column_indexes():
    header_row_idx = 1
    for each_key in column_idx_dict.keys():
        print(each_key, ':' ,source_worksheet[header_row_idx][column_idx_dict[each_key]].value)

def check_closed_conditions(current_row):
    return_val = False
    if current_row[column_idx_dict['close_week']].value.strip() != '':
        if int(current_row[column_idx_dict['close_week']].value.strip()) == int(constant_dict['file_name'][4:6]):
            if int(current_row[column_idx_dict['close_period']].value.strip()) == int(constant_dict['file_name'][1:3]):
                if int(current_row[column_idx_dict['close_year']].value.strip()) == datetime.datetime.now().year:
                    return_val = True
    return return_val

def check_open_conditions(current_row):
    return_val = False
    if current_row[column_idx_dict['close_week']].value.strip() == '':
        if int(current_row[column_idx_dict['open_week']].value.strip()) == int(constant_dict['file_name'][4:6]):
            if int(current_row[column_idx_dict['open_period']].value.strip()) == int(constant_dict['file_name'][1:3]):
                if int(current_row[column_idx_dict['open_year']].value.strip()) == datetime.datetime.now().year:
                    return_val = True
    return return_val


def extract_data():
    result_data_list = []
    for row_idx in range(1, (source_worksheet.max_row + 1)):
        if source_worksheet[row_idx][column_idx_dict['assignment_team']].value == constant_dict['assignment_team_value']:
            if check_closed_conditions(source_worksheet[row_idx]) or check_open_conditions(source_worksheet[row_idx]):
                result_data_list.append(source_worksheet[row_idx])
    return result_data_list

print(len(extract_data()))
