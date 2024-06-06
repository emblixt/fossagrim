import os.path

from pathlib import Path
import openpyxl

import Fossagrim.io.fossagrim_io as fio
from Fossagrim.utils.definitions import heureka_standdata_keys, heureka_standdata_desc


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
    file_list = Path("C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger").rglob('*result*xlsx')
    for result_file in file_list:
        pass
        # rename_avg_stands(result_file)


def collect_all_stand_data():
    base_dir = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger"
    result_file = os.path.join(base_dir, 'All stand data.csv')
    # create header lines of result file
    fio.write_csv_file(result_file, heureka_standdata_keys, heureka_standdata_desc)

    file_list = Path(base_dir).rglob('*Averaged stand data.csv')
    with open(result_file, 'a') as _out:  # append
        for result_file in file_list:
            print(result_file)
            with open(result_file, 'r') as _in:
                for line in _in.readlines():
                    if 'FHF' in line[:10]:
                        _out.write(line)



def test_rename_avg_stands():
    f = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0016 Margrete Folsland\\FHF24-0016 Heureka results.xlsx"
    rename_avg_stands(f)


if __name__ == '__main__':
    collect_all_stand_data()