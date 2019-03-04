import openpyxl
import argparse
import os
import sys
import json


filepath = os.path.join(sys.path[0], "purch_work.json")


def load_json_data(filepath):
    if not os.path.exists(filepath):
        return None
    with open(filepath, 'r') as file_handler:
        return json.load(file_handler)

def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('file_with_address', help='A file which contains addresses')
    arg = parser.parse_args()
    return arg.file_with_address




def write_infromation_into_file(my_file, information_from_json):
    work_book = openpyxl.load_workbook(filename=my_file)
    work_sheet = work_book.active
    work_sheet['B4'] = information_from_json['principal']['full_name']
    work_sheet['D18'] = information_from_json['principal']['short_name']
    work_sheet['D19'] = information_from_json['principal']['inn']
    work_sheet['D20'] = information_from_json['principal']['kpp']
    work_sheet['D21'] = information_from_json['principal']['ogrn']
    work_book.save(my_file)


if __name__ == '__main__':
    file_name = get_args()
    information_from_json = load_json_data(filepath)
    write_infromation_into_file(file_name, information_from_json)