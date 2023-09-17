import os
import numpy as np
import pandas as pd
import openpyxl
from copy import deepcopy
import Fossagrim.io.fossagrim_io as fio

methods = ['BAU', 'PRES']  # "Business as usual" and "Preservation"


def arrange_import(_stand_file, _csv_stand_file, _csv_treatment_file, _average_over, _stand_id_key, _project_tag,
                   verbose=False):
    fio.export_fossagrim_stand_to_heureka(
        _stand_file,
        _csv_stand_file,
        average_over=_average_over,
        stand_id_key=_stand_id_key,
        average_name='{} Avg Stand'.format(_project_tag),
        verbose=verbose
    )

    fio.export_fossagrim_treatment(
        _stand_file,
        _csv_treatment_file,
        average_over=_average_over,
        stand_id_key=_stand_id_key,
        average_name='{} Avg Stand'.format(_project_tag)
    )


def arrange_results(_result_file, _sheet_names, _combine_sheets, _monetization_file):
    # sheet names in the resulting excel file
    write_monetization_file = \
        fio.rearrange_raw_heureka_results(
            _result_file, _sheet_names, _combine_sheets, monetization_file=_monetization_file)

    if write_monetization_file:
        print('Try to modify Monetization file')
        # TODO
        # implement this
        fio.modify_monetization_file(_monetization_file)


def project_settings(_project_tag, _project_settings_file):
    """
    Reads the project settings excel file, and extracts information from it to be used in the Heureka simulation
    and sets up files necessary for storing results and monetization calculations for the given project

    Keys expected in project settings files:
    "Project name", "Status", "Forest types", "Average over stands", "Position of fractions when combined",
    "Check length of lists", "Stand id key", "Stands file", "Results file", "Monetization file"

    :param _project_tag:
        str
        Name tag that identifies the current project, e.g. "FHF23-007"

    :param _project_settings_file:
        str
        Full path name of the project settings excel sheet.
    :return:
    """
    p_tabl = fio.read_excel(_project_settings_file, 1, 'Settings')
    w_dir = os.path.dirname(_project_settings_file)

    i = None
    for i, p_name in enumerate(p_tabl['Project name']):
        if _project_tag in p_name:
            break

    _project_folder = os.path.join(w_dir, p_tabl['Project name'][i])
    _stand_id_key = p_tabl['Stand id key'][i]
    _stand_file = os.path.join(_project_folder, p_tabl['Stands file'][i])
    _result_file = os.path.join(_project_folder, p_tabl['Results file'][i])
    _monetization_file = os.path.join(_project_folder, p_tabl['Monetization file'][i])
    _csv_stand_file = os.path.join(_project_folder, '{} Averaged stand data.csv'.format(_project_tag))
    _csv_treatment_file = os.path.join(_project_folder, '{} Averaged treatment.csv'.format(_project_tag))

    forest_types = [_x.strip() for _x in p_tabl['Forest types'][i].split(',')]
    average_over_strings = [_x.strip() for _x in p_tabl['Average over stands'][i].split(',')]
    if len(forest_types) != len(average_over_strings):
        raise ValueError('Number of forest types ({}) is not the same as Average stands ({})'.format(
            len(forest_types), len(average_over_strings)))
    _average_over = {}
    for j, f_type in enumerate(forest_types):
        if _stand_id_key == 'Bestand':
            _average_over[f_type] = [int(_x.strip()) for _x in average_over_strings[j].split(';')]
        else:
            _average_over[f_type] = [_x.strip() for _x in average_over_strings[j].split(';')]

    _result_sheets = \
        ['{} Avg Stand-{} {}'.format(_project_tag, _x, _y) for _x, _y in
         zip(np.repeat(list(_average_over.keys()), len(_average_over)), methods * len(_average_over))]

    # Create empty results file
    if os.path.isfile(_result_file):
        print("WARNING result file {} already exists. No empty result file created".format(
            os.path.split(_result_file)[-1]))
    else:
        writer = pd.ExcelWriter(_result_file, engine='xlsxwriter')
        wb = writer.book
        for _sheet in _result_sheets:
            ws = wb.add_worksheet(_sheet)
        writer.close()

    combine_positions = [_x.strip() for _x in p_tabl['Position of fractions when combined'][i].split(',')]
    wb = openpyxl.load_workbook(_stand_file, data_only=True)
    ws = wb.active
    combine_fractions = [ws[_x].value for _x in combine_positions]
    wb.close()
    _combine_sheets = {}
    for j, method in enumerate(methods):
        _this_list = []
        for k, c_frac in enumerate(combine_fractions):
            _this_list.append(_result_sheets[j + 2 * k])
            _this_list.append(c_frac)
        _combine_sheets['{} Combined Stands {} {}'.format(_project_tag, '-and-'.join(forest_types), method)] = \
            _this_list

    return _project_folder, _stand_file, _average_over, _stand_id_key, _result_file, _result_sheets, _combine_sheets, \
        _monetization_file, _csv_stand_file, _csv_treatment_file


def project_settings_old(_project_tag):

    # ################# FHF23-003 ######################################################################################
    if _project_tag == 'FHF23-003':
        _project_folder = \
            "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\{} Kloppmyra".format(_project_tag)
        _average_over = \
            {'Spruce': [67, 68, 70, 72, 75, 79, 80, 85, 88, 90, 92, 93, 95, 97, 98, 99, 104, 105, 109]}
        _stand_id_key = 'Bestand'
        _result_sheets = [
            '{} Business as usual'.format(_project_tag), '{} Preservation'.format(_project_tag)
        ]
        _combine_sheets = None
        _monet_file = os.path.join(_project_folder, '{} Monetization_V2 WIP.xlsx'.format(_project_tag))
        _monet_file_header_lines = 5

    # ################# FHF23-004 ######################################################################################
    elif _project_tag == 'FHF23-004':
        _project_folder = \
            "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\{} Høgberget".format(_project_tag)
        _average_over = {
            'Spruce': [5, 12],
            'Pine': [1, 7, 8, 9, 13, 16, 17]
        }
        _stand_id_key = 'Bestand'
        _result_sheets = [
            '{} Avg Stand-Spruce BAU'.format(_project_tag), '{} Avg Stand-Spruce PRES'.format(_project_tag),
            '{} Avg Stand-Pine BAU'.format(_project_tag), '{} Avg Stand-Pine PRES'.format(_project_tag)
        ]
        # Combine Heureka simulation results from different forest types.
        # The key will be the name of the combination and the value, the list, consists of the sheet name from where
        # to pick up the results for this forest type, and the excel column - row index from where to pick up the
        # fraction this forest type makes
        _combine_sheets = {
            '{} Combined Stands spruce-and-pine BAU'.format(_project_tag):
                [_result_sheets[0], 'BD6', _result_sheets[2], 'BE6'],
            '{} Combined Stands spruce-and-pine PRES'.format(_project_tag):
                [_result_sheets[1], 'BD6', _result_sheets[3], 'BE6']
        }

    # ################# FHF23-005 ######################################################################################
    elif _project_tag == 'FHF23-005':
        _project_folder = \
            "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\{} Kvistaul".format(_project_tag)
        _average_over = {
            'Spruce': ['FHF23-005-{}'.format(_x) for _x in
                       ['1', '2', '3A', '3B', '4', '11', '13', '14', '15', '19', '28',
                        '30', '31', '32', '33', '37', '38', '39', '41', '23']],
            'Pine': ['FHF23-005-{}'.format(_x) for _x in [18, 26, 27, 29]]
        }
        _stand_id_key = 'Fossagrim ID'
        _result_sheets = [
            '{} Avg Stand-Spruce BAU'.format(_project_tag), '{} Avg Stand-Spruce PRES'.format(_project_tag),
            '{} Avg Stand-Pine BAU'.format(_project_tag), '{} Avg Stand-Pine PRES'.format(_project_tag)
        ]
        _combine_sheets = None

    # ################# FHF23-006 ######################################################################################
    elif _project_tag == 'FHF23-006':
        _project_folder = \
            "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\{} Sondre Øverbø".format(_project_tag)
        _average_over = {
            'Spruce': ['FHF23-006-{}'.format(_x) for _x in
                       ['1', '3A', '3B', '33', '48', '54', '57']],
            'Pine': ['FHF23-006-38']
        }
        _stand_id_key = 'Fossagrim ID'
        _result_sheets = [
            '{} Avg Stand-Spruce BAU'.format(_project_tag), '{} Avg Stand-Spruce PRES'.format(_project_tag),
            '{} Avg Stand-Pine BAU'.format(_project_tag), '{} Avg Stand-Pine PRES'.format(_project_tag)
            ]
        _combine_sheets = None

    # ################# FHF23-007 ######################################################################################
    elif _project_tag == 'FHF23-007':
        _project_folder = \
            "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\{} Gudmund Aaker".format(_project_tag)
        _average_over = {
            'Spruce': [1, 4, 8],
            'Pine': [2, 3, 5, 6, 7]
        }
        _stand_id_key = 'Bestand'
        _result_sheets = [
            '{} Avg Stand-Spruce BAU'.format(_project_tag), '{} Avg Stand-Spruce PRES'.format(_project_tag),
            '{} Avg Stand-Pine BAU'.format(_project_tag), '{} Avg Stand-Pine PRES'.format(_project_tag)
        ]
        _combine_sheets = None

    # ################# FHF23-008 ######################################################################################
    elif _project_tag == 'FHF23-008':
        _project_folder = \
            "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\{} Ola Føsker".format(_project_tag)
        _average_over = {
            'Birch': [85],
            'Pine': [27, 22, 21]
        }
        _stand_id_key = 'Bestand'
        _result_sheets = [
            '{} Avg Stand-Birch BAU'.format(_project_tag), '{} Avg Stand-Birch PRES'.format(_project_tag),
            '{} Avg Stand-Pine BAU'.format(_project_tag), '{} Avg Stand-Pine PRES'.format(_project_tag)
        ]
        _combine_sheets = None

    # ################# FHF23-009 ######################################################################################
    elif _project_tag == 'FHF23-009':
        _project_folder = \
            "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\{} Galterud".format(_project_tag)
        _average_over = {
            'Spruce': [16, 101],
            'Pine': [3, 8, 13, 17, 96]
        }
        _stand_id_key = 'Bestand'
        _result_sheets = \
            ['{} Avg Stand-{} {}'.format(_project_tag, _x, _y) for _x, _y in
             zip(np.repeat(list(_average_over.keys()), len(_average_over)), methods * len(_average_over))]
        # _result_sheets = [
        #     '{} Avg Stand-Spruce BAU'.format(_project_tag), '{} Avg Stand-Spruce PRES'.format(_project_tag),
        #     '{} Avg Stand-Pine BAU'.format(_project_tag), '{} Avg Stand-Pine PRES'.format(_project_tag)
        # ]
        _combine_sheets = {
            '{} Combined Stands {} BAU'.format(_project_tag, '-and-'.join(list(_average_over.keys()))):
                [_result_sheets[0], 'BD6', _result_sheets[2], 'BE6'],
            '{} Combined Stands {} PRES'.format(_project_tag, '-and-'.join(list(_average_over.keys()))):
                [_result_sheets[1], 'BD6', _result_sheets[3], 'BE6']
        }

    # ################# FHF23-010 ######################################################################################
    elif _project_tag == 'FHF23-010':
        _project_folder = \
            "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\{} Heimyra".format(_project_tag)
        _average_over = {
            'Spruce': [45, 54, 38],
            'Pine': [36, 35]
        }
        _stand_id_key = 'Bestand'
        _result_sheets = [
            '{} Avg Stand-Spruce BAU'.format(_project_tag), '{} Avg Stand-Spruce PRES'.format(_project_tag),
            '{} Avg Stand-Pine BAU'.format(_project_tag), '{} Avg Stand-Pine PRES'.format(_project_tag)
        ]
        _combine_sheets = {
            '{} Combined Stands spruce-and-pine BAU'.format(_project_tag):
                [_result_sheets[0], 'BD6', _result_sheets[2], 'BE6'],
            '{} Combined Stands spruce-and-pine PRES'.format(_project_tag):
                [_result_sheets[1], 'BD6', _result_sheets[3], 'BE6']
        }

    # ################# FHF23-011 ######################################################################################
    elif _project_tag == 'FHF23-011':
        _project_folder = \
            "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\{} Ola Haukerud_Styggdalen".format(
                _project_tag)
        _average_over = {
            'Spruce': [5, 14],
            'Pine': [3, 4, 6, 7, 8, 9, 11, 13]
        }
        _stand_id_key = 'Bestand'
        _result_sheets = [
            '{} Avg Stand-Spruce BAU'.format(_project_tag), '{} Avg Stand-Spruce PRES'.format(_project_tag),
            '{} Avg Stand-Pine BAU'.format(_project_tag), '{} Avg Stand-Pine PRES'.format(_project_tag)
        ]
        _combine_sheets = {
            '{} Combined Stands spruce-and-pine BAU'.format(_project_tag):
                [_result_sheets[0], 'BD6', _result_sheets[2], 'BE6'],
            '{} Combined Stands spruce-and-pine PRES'.format(_project_tag):
                [_result_sheets[1], 'BD6', _result_sheets[3], 'BE6']
        }

    # ################# FHF23-012 ######################################################################################
    elif _project_tag == 'FHF23-012':
        _project_folder = \
            "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\{} Ola Haukerud_Vestskogen".format(
                _project_tag)
        _average_over = {
            'Spruce': [1, 3, 14, 19],
            'Pine': [2, 18, 21, 22, 24, 25, 27]
        }
        _stand_id_key = 'Bestand'
        _result_sheets = [
            '{} Avg Stand-Spruce BAU'.format(_project_tag), '{} Avg Stand-Spruce PRES'.format(_project_tag),
            '{} Avg Stand-Pine BAU'.format(_project_tag), '{} Avg Stand-Pine PRES'.format(_project_tag)
        ]
        _combine_sheets = {
            '{} Combined Stands spruce-and-pine BAU'.format(_project_tag):
                [_result_sheets[0], 'BD6', _result_sheets[2], 'BE6'],
            '{} Combined Stands spruce-and-pine PRES'.format(_project_tag):
                [_result_sheets[1], 'BD6', _result_sheets[3], 'BE6']
        }
    else:
        raise IOError('Project {} is not known'.format(_project_tag))

    _stand_file = os.path.join(_project_folder, '{} Bestandsutvalg.xlsx'.format(_project_tag))
    _result_file = os.path.join(_project_folder, '{} Heureka results.xlsx'.format(_project_tag))
    _csv_stand_file = os.path.join(_project_folder, '{} Averaged stand data.csv'.format(_project_tag))
    _csv_treatment_file = os.path.join(_project_folder, '{} Averaged treatment.csv'.format(_project_tag))

    # Create empty results file
    if os.path.isfile(_result_file):
        print("WARNING result file {} already exists. No empty result file created".format(
            os.path.split(_result_file)[-1]))
    else:
        writer = pd.ExcelWriter(_result_file, engine='xlsxwriter')
        wb = writer.book
        for _sheet in _result_sheets:
            ws = wb.add_worksheet(_sheet)
        writer.close()

    # For Combined calculations (e.g. combining the results from both Spruce and Pine) extract the fractions of the
    # different forest types and replace the excel positions with this fraction
    if isinstance(_combine_sheets, dict):
        wb = openpyxl.load_workbook(_stand_file, data_only=True)
        ws = wb.active
        for key in list(_combine_sheets.keys()):
            this_list = deepcopy(_combine_sheets[key])
            for i, value in enumerate(_combine_sheets[key]):
                if np.mod(i, 2) == 1:  # the excel positions must come in on every second item in this list
                    this_fraction = ws[_combine_sheets[key][i]].value
                    this_list[i] = this_fraction
            _combine_sheets[key] = this_list
        wb.close()

    return _project_folder, _stand_file, _average_over, _stand_id_key, _result_file, _result_sheets, _combine_sheets, \
        _csv_stand_file, _csv_treatment_file


if __name__ == '__main__':
    project_tag = 'FHF23-007'
    project_settings_file = 'C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\ProjectForestsSettings.xlsx'

    fix_import = False  # Set to False after Heureka simulation results have been saved in result_file and you want to
                        # rearrange the results so that they are easier to include in excel calculations

    verbose = False

    # project_folder, stand_file, average_over, stand_id_key, result_file, result_sheets, combine_sheets, \
    # csv_stand_file, csv_treatment_file = project_settings_old(project_tag)
    project_folder, stand_file, average_over, stand_id_key, result_file, result_sheets, combine_sheets, \
        monetization_file, csv_stand_file, csv_treatment_file = project_settings(project_tag, project_settings_file)

    if fix_import:
        arrange_import(stand_file, csv_stand_file, csv_treatment_file, average_over, stand_id_key, project_tag,
                       verbose=verbose)
    else:
        arrange_results(result_file, result_sheets, combine_sheets, monetization_file)

