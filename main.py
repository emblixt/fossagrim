import os
import pandas as pd
import Fossagrim.io.fossagrim_io as fio


def arrange_tags_and_files(_project_tag, _project_folder, _result_file, _result_sheets):
    _stand_file = os.path.join(_project_folder, '{} Bestandsutvalg.xlsx'.format(_project_tag))

    _csv_stand_file = os.path.join(_project_folder, '{} Averaged stand data.csv'.format(_project_tag))
    _csv_treatment_file = os.path.join(_project_folder, '{} Averaged treatment.csv'.format(_project_tag))

    # Create empty results file
    if os.path.isfile(_result_file):
        print("WARNING result file {} already exists. No empty result file created".format(os.path.split(_result_file)[-1]))
    else:
        writer = pd.ExcelWriter(_result_file, engine='xlsxwriter')
        wb = writer.book
        for _sheet in _result_sheets:
            ws = wb.add_worksheet(_sheet)
        writer.close()

    return _stand_file, _csv_stand_file, _csv_treatment_file


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


def arrange_results(_result_file, _sheet_names):
    # sheet names in the resulting excel file
    fio.rearrange_raw_heureka_results(_result_file, _sheet_names)


def project_settings(_project_tag):

    # ################# FHF23-003 ######################################################################################
    if _project_tag == 'FHF23-003':
        _project_folder = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-003 Kloppmyra"
        _average_over = \
            {'Spruce': [67, 68, 70, 72, 75, 79, 80, 85, 88, 90, 92, 93, 95, 97, 98, 99, 104, 105, 109]}
        _stand_id_key = 'Bestand'
        _result_file = os.path.join(_project_folder, '{} Heureka results.xlsx'.format(_project_tag))
        _result_sheets = [
            'FHF23-003 Business as usual', 'FHF23-003 Preservation'
        ]

    # ################# FHF23-004 ######################################################################################
    elif _project_tag == 'FHF23-004':
        _project_folder = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-004 Høgberget"
        _average_over = {
            'Spruce': [5, 12],
            'Pine': [1, 7, 8, 9, 13, 16, 17]
        }
        _stand_id_key = 'Bestand'
        _result_file = os.path.join(_project_folder, '{} Heureka results.xlsx'.format(_project_tag))
        _result_sheets = [
            'FHF23-004 Avg Stand-Spruce BAU', 'FHF23-004 Avg Stand-Spruce PRES',
            'FHF23-004 Avg Stand-Pine BAU', 'FHF23-004 Avg Stand-Pine PRES'
        ]

    # ################# FHF23-005 ######################################################################################
    elif _project_tag == 'FHF23-005':
        _project_folder = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-005 Kvistaul"
        _average_over = {
            'Spruce': ['FHF23-005-{}'.format(_x) for _x in
                       ['1', '2', '3A', '3B', '4', '11', '13', '14', '15', '19', '28',
                        '30', '31', '32', '33', '37', '38', '39', '41', '23']],
            'Pine': ['FHF23-005-{}'.format(_x) for _x in [18, 26, 27, 29]]
        }
        _stand_id_key = 'Fossagrim ID'
        _result_file = os.path.join(_project_folder, '{} Heureka results.xlsx'.format(_project_tag))
        _result_sheets = [
            'FHF23-005 Avg Stand-Spruce BAU', 'FHF23-005 Avg Stand-Spruce PRES',
            'FHF23-005 Avg Stand-Pine BAU', 'FHF23-005 Avg Stand-Pine PRES'
        ]

    # ################# FHF23-006 ######################################################################################
    elif _project_tag == 'FHF23-006':
        _project_folder = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-006 Sondre Øverbø"
        _average_over = {
            'Spruce': ['FHF23-006-{}'.format(_x) for _x in
                       ['1', '3A', '3B', '33', '48', '54', '57']],
            'Pine': ['FHF23-006-38']
        }
        _stand_id_key = 'Fossagrim ID'
        _result_file = os.path.join(_project_folder, '{} Heureka results.xlsx'.format(_project_tag))
        _result_sheets = [
            'FHF23-006 Avg Stand-Spruce BAU', 'FHF23-006 Avg Stand-Spruce PRES',
            'FHF23-006 Avg Stand-Pine BAU', 'FHF23-006 Avg Stand-Pine PRES'
            ]

    # ################# FHF23-007 ######################################################################################
    elif _project_tag == 'FHF23-007':
        _project_folder = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-007 Gudmund Aaker"
        _average_over = {
            'Spruce': [1, 4, 8],
            'Pine': [2, 3, 5, 6, 7]
        }
        _stand_id_key = 'Bestand'
        _result_file = os.path.join(_project_folder, '{} Heureka results.xlsx'.format(_project_tag))
        _result_sheets = [
            'FHF23-007 Avg Stand-Spruce BAU', 'FHF23-007 Avg Stand-Spruce PRES',
            'FHF23-007 Avg Stand-Pine BAU', 'FHF23-007 Avg Stand-Pine PRES'
        ]

    # ################# FHF23-008 ######################################################################################
    elif _project_tag == 'FHF23-008':
        _project_folder = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-008 Ola Føsker"
        _average_over = {
            'Birch': [85],
            'Pine': [27, 22, 21]
        }
        _stand_id_key = 'Bestand'
        _result_file = os.path.join(_project_folder, '{} Heureka results.xlsx'.format(_project_tag))
        _result_sheets = [
            'FHF23-008 Avg Stand-Birch BAU', 'FHF23-008 Avg Stand-Birch PRES',
            'FHF23-008 Avg Stand-Pine BAU', 'FHF23-008 Avg Stand-Pine PRES'
        ]

    # ################# FHF23-009 ######################################################################################
    elif _project_tag == 'FHF23-009':
        _project_folder = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-009 Galterud"
        _average_over = {
            'Pine': [3, 8, 13, 17]
        }
        _stand_id_key = 'Bestand'
        _result_file = os.path.join(_project_folder, '{} Heureka results.xlsx'.format(_project_tag))
        _result_sheets = [
            'FHF23-009 Avg Stand-Pine BAU', 'FHF23-009 Avg Stand-Pine PRES'
        ]

    else:
        raise IOError('Project {} is not known'.format(_project_tag))

    return _project_folder, _average_over, _stand_id_key, _result_file, _result_sheets


if __name__ == '__main__':
    project_tag = 'FHF23-009'
    fix_import = False

    verbose = False

    project_folder, average_over, stand_id_key, result_file, result_sheets = project_settings(project_tag)

    if fix_import:
        stand_file, csv_stand_file, csv_treatment_file = \
            arrange_tags_and_files(project_tag, project_folder, result_file, result_sheets)

        arrange_import(stand_file, csv_stand_file, csv_treatment_file, average_over, stand_id_key, project_tag,
                       verbose=verbose)
    else:
        arrange_results(result_file, result_sheets)

