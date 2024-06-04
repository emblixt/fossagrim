import os.path

import openpyxl

import Fossagrim.io.fossagrim_io as fio


def rename_avg_stands(_result_file):
    repl_str = 'Avg Stand-'
    print('Working on {}'.format(os.path.basename(_result_file)))
    wb = openpyxl.load_workbook(_result_file)
    for sheet_name in wb.sheetnames:
        if repl_str in sheet_name:
            print(' - ')
            print(' - Sheet: {}'.format(sheet_name))
            new_sheet_name = sheet_name.replace(repl_str, '')
            this_sheet = wb[sheet_name]
            this_sheet.title = new_sheet_name
            print(' - Renamed to: {}'.format(new_sheet_name))
    wb.save(_result_file)


def find_and_rename_avg_stands():
    from pathlib import Path
    file_list = Path("C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger").rglob('*result*xlsx')
    for result_file in file_list:
        # rename_avg_stands(result_file)


def test_rename_avg_stands():
    f = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0016 Margrete Folsland\\FHF24-0016 Heureka results.xlsx"
    rename_avg_stands(f)


if __name__ == '__main__':
    find_and_rename_avg_stands()