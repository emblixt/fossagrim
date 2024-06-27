import os.path

from pathlib import Path

import matplotlib.pyplot as plt
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


def collect_carbon_effect_for_test_simulations(stand_file, result_file, out_file, postfix=None):
    """

    :param stand_file:
        str
        Full path name of csv stand file which was used as input to the Heureka simulations
    :param result_file:
        str
        Full path name to excel file which contain the raw Heureka results from above simulation
    :param out_file:
        str
        Full path name of csv stand file that is a copy of the input stand file, but with the
        calculated 30 year carbon effect added to it
        If file exists, it appends to it
    :param postfix:
        str
        Optional
        String that is added to each stand ID to separate it from other earlier stand ID's
    :return:
    """
    # test if output file exists
    if os.path.isfile(out_file):
        append = True
    else:
        append = False
        # create header lines of out file
        fio.write_csv_file(out_file, heureka_standdata_keys, heureka_standdata_desc)
    append = True

    fig, ax = plt.subplots()

    # Open output csv file for append
    with open(out_file, 'a') as _out:
        # Open stand file to read each stand
        with open(stand_file, 'r') as _in:
            for i, line in enumerate(_in.readlines()):
                if i < 2:
                    continue
                split_line = line.split(';')
                print('Working on stand {}'.format(split_line[0]))
                carbon_effect = fio.get_carbon_effect(result_file, split_line[0], ax=ax)
                if carbon_effect is None:
                    print('  Carbon effect: None')
                    carbon_effect = ''
                else:
                    print('  Carbon effect: {:.2f}'.format(carbon_effect))

                split_line[101] = str(carbon_effect)
                _out.write(';'.join(split_line))


def test_rename_avg_stands():
    f = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0016 Margrete Folsland\\FHF24-0016 Heureka results.xlsx"
    rename_avg_stands(f)


if __name__ == '__main__':
    # collect_all_stand_data()
    # correct_spelling()
    # merge_treatment_and_stand()
    stand_file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\TestingHeurekaParameters\\Stands and treatments Pine 2.csv"
    result_file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\TestingHeurekaParameters\\Heureka results Pine 2.xlsx"
    out_file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\TestingHeurekaParameters\\Test stands with carbon effect.csv"
    # collect_carbon_effect_for_test_simulations(stand_file, result_file, out_file)
    plt.show()