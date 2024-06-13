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


def correct_spelling():
    import shutil
    base_dir = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger"
    file_list = Path(base_dir).rglob('*Averaged treatment.csv')
    for file in file_list:
        # create a backup file
        orig_name = str(file)
        new_name = orig_name.replace('.csv', '_BACKUP.csv')
        print(new_name)
        shutil.copy(orig_name, new_name)
        with open(new_name) as input_file:
            with open(orig_name, 'w') as output_file:
                lines = input_file.readlines()
                for line in lines:
                    output_file.write(line.replace('Feeling', 'Felling'))


def merge_treatment_and_stand():
    """
    Adds the treatment to the stand file as UserDefinedVariables
    :return:
    """
    import shutil
    base_dir = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger"
    treatment_files = Path(base_dir).rglob('*Averaged treatment.csv')
    for i, tf in enumerate(treatment_files):
        # stand_id = tf.name[:10].strip()
        stand_id = tf.name.split()[0].strip()
        if stand_id == 'FHF23-999':  # test data
            continue
        stand_files = Path(base_dir).rglob('*Averaged stand data.csv')
        for j, sf in enumerate(stand_files):
            if sf.name.replace('Averaged stand data', 'Averaged treatment') == tf.name:
                print(stand_id, tf.name, sf.name)
                orig_name = str(sf)
                new_name = orig_name.replace('.csv', '_BACKUP.csv')
                shutil.copy(orig_name, new_name)
                print('write header lines to {}'.format(os.path.basename(orig_name)))
                fio.write_csv_file(orig_name, heureka_standdata_keys, heureka_standdata_desc)
                with open(orig_name, 'a') as oid:
                    print(' Open {} for appending'.format(os.path.basename(orig_name)))
                    with open(new_name, 'r') as fid:
                        print('  Open {} for reading'.format(os.path.basename(new_name)))
                        lines = fid.readlines()
                        # open file and write
                        for line in lines:
                            for ws in ['Spruce', 'Pine', 'Birch', 'Avg Stand-Spruce', 'Avg Stand-Pine', 'Avg Stand-Birch']:
                                # print('   {}'.format(ws))
                                treatment = fio.read_fossagrim_treatment(tf, stand_id, ws)
                                if None in treatment:
                                    continue
                                if '{} {}'.format(stand_id, ws) in line:
                                    orig_line = line.split(';')
                                    new_line = orig_line
                                    new_line[98] = str(treatment[1])
                                    new_line[99] = str(treatment[0])
                                    new_line[100] = str(treatment[2])
                                    print('    - Add treatment to {} {}'.format(stand_id, ws) )
                                    oid.write(';'.join(new_line))



def test_rename_avg_stands():
    f = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0016 Margrete Folsland\\FHF24-0016 Heureka results.xlsx"
    rename_avg_stands(f)


if __name__ == '__main__':
    # collect_all_stand_data()
    # correct_spelling()
    merge_treatment_and_stand()