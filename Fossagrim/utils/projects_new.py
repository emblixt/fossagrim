import os
import numpy as np
import pandas as pd
import openpyxl
# from copy import deepcopy
import Fossagrim.plotting.misc_plots as fpp
import Fossagrim.io.fossagrim_io as fio

methods = ['BAU', 'PRES']  # "Business as usual" and "Preservation"


def arrange_import(_stand_file, _csv_stand_file, _csv_treatment_file, _average_over, _stand_id_key, _project_tag,
                   _verbose=False, **_kwargs):

    # QC input table
    table = fio.read_excel(_stand_file, 7, 0)
    fpp.qc_stand_data(
        table,
        os.path.basename(_stand_file),
        os.path.join(os.path.dirname(_stand_file), 'QC_plots')
    )

    fio.export_fossagrim_stand_to_heureka(
        _stand_file,
        _csv_stand_file,
        average_over=_average_over,
        stand_id_key=_stand_id_key,
        average_name='{} '.format(_project_tag),
        verbose=_verbose
    )

    # TODO
    # Add the treatment to the csv stand file instead, among the UserDefinedVariables, then we don't need this file
    fio.export_fossagrim_treatment(
        _stand_file,
        _csv_treatment_file,
        average_over=_average_over,
        stand_id_key=_stand_id_key,
        average_name='{} '.format(_project_tag)
    )


def arrange_results(_result_file, _sheet_names, _combine_sheets, _monetization_file, _verbose=False, **_kwargs):
    # sheet names in the resulting Excel file
    write_monetization_file = \
        fio.rearrange_raw_heureka_results(
            _result_file, _sheet_names, _combine_sheets, monetization_file=_monetization_file, verbose=_verbose)

    if write_monetization_file:
        print('Try to modify Monetization file')
        fio.modify_monetization_file(_monetization_file, **_kwargs)


def project_settings(_project_name, _project_settings_file, _fix_import: bool = True):
    """
    Reads the project settings Excel file, and extracts information from it to be used in the Heureka simulation
    and sets up files necessary for storing results and monetization calculations for the given project

    Keys expected in project settings files:
    "Project name", "Project folder", "Status", "Stand id key", "Stands file", "Results file",

    :param _project_name:
        str
        Name that identifies the current project, e.g. "FHF23-007 Gudmund Aaker"

    :param _project_settings_file:
        str
        Full path name of the project settings Excel sheet.

    :param _fix_import:
        Bool
        If True it goes through the whole process and creates results files,
        If False, it only returns the first few arguments before creating any files and thereby
        avoiding the possibility of overwriting the results file.
    :return:
    """
    p_tabl = fio.read_excel(_project_settings_file, 1, 'Settings')

    i = None
    tag_found = False
    for i, p_name in enumerate(p_tabl['Project name']):
        if not isinstance(p_name, str):
            continue
        if p_tabl['Status'][i] not in ['Active', 'Prospective']:
            continue
        if _project_name in p_name.strip():
            tag_found = True
            break

    if not tag_found:
        raise IOError('Project name {} not found in {}'.format(
            _project_name, os.path.basename(_project_settings_file)
        ))

    _project_folder = p_tabl['Project folder'][i]
    _qc_folder = os.path.join(_project_folder, 'QC_plots')
    _stand_id_key = p_tabl['Stand id key'][i]
    _stand_file = os.path.join(_project_folder, p_tabl['Stands file'][i])
    _result_file = os.path.join(_project_folder, p_tabl['Results file'][i])
    # _monetization_file = os.path.join(_project_folder, p_tabl['Monetization file'][i])  # not calculated in Python
    _csv_stand_file = os.path.join(_project_folder, '{} Averaged stand data.csv'.format(_project_name))
    _csv_treatment_file = os.path.join(_project_folder, '{} Averaged treatment.csv'.format(_project_name))

    _average_over, _active_prod_areas, _active_prod_areas_fract = fio.get_active_forests_from_stand(
        _stand_file, stand_id_key=_stand_id_key)

    _result_sheets = \
        ['{} {} {}'.format(_project_name, _x, _y) for _x, _y in
         zip(np.repeat(list(_average_over.keys()), 2), methods * len(_average_over))]
         # zip(np.repeat(list(_average_over.keys()), max(2, len(_average_over))), methods * len(_average_over))]

    # We could extract the combine fractions directly from the stand file (Bestandsutvalg,
    # through "fio.get_active_forests_from_stand()" above) because it should
    # contain the areas/volumes of each of the different wood species
    for key, value in _active_prod_areas.items():
        print('{} stands have a total area of {} daa, which corresponds to a fraction of {:.3} of the total area'.format(
            key, sum(value), _active_prod_areas_fract[key]))

    _combine_sheets = {}
    # _kwargs, combine_fractions = fio.get_kwargs_from_stand(_stand_file, _project_settings_file, _project_tag)
    # forest_types = [_x.strip() for _x in p_tabl['Forest types'][i].split(',')] # From OLD projects.py

    # With the new set-up (reading active productive area and fractions directly from Bestandsutvalg file)
    combine_fractions = list(_active_prod_areas_fract.values())
    forest_types = list(_active_prod_areas_fract.keys())

    if len(combine_fractions) < 2:
        # only one forest type to "combine", so no combination of forest types necessary,
        # and the 'combine_fraction' is overridden using a factor of 1
        _combine_fractions = [1.]
    else:
        _combine_fractions = combine_fractions

    for j, method in enumerate(methods):
        _this_list = []
        for k, c_frac in enumerate(_combine_fractions):
            _this_list.append(_result_sheets[j + 2 * k])
            _this_list.append(c_frac)
        _combine_sheets['{} Combined Stands {} {}'.format(_project_name, '-and-'.join(forest_types), method)] = \
            _this_list

    if not _fix_import:
        return (_project_folder, _stand_file, _average_over, _stand_id_key, _result_file, _result_sheets,
                _combine_sheets, None, _csv_stand_file, _csv_treatment_file, {})

    # Create empty QC folder if it doesn't exist from before
    if not os.path.exists(_qc_folder):
        os.makedirs(_qc_folder)

    # Create empty results file
    create_results_file = False
    if os.path.isfile(_result_file):
        print("WARNING result file {} already exists.".format(
            os.path.split(_result_file)[-1]))
        _ans = input("Do you want to overwrite? Y/[N]:")
        if _ans.upper() == 'Y':
            create_results_file = True
    else:
        create_results_file = True

    if create_results_file:
        writer = pd.ExcelWriter(_result_file, engine='xlsxwriter')
        wb = writer.book
        for _sheet in _result_sheets:
            _ = wb.add_worksheet(_sheet)
        writer.close()

    return _project_folder, _stand_file, _average_over, _stand_id_key, _result_file, _result_sheets, _combine_sheets, \
        None, _csv_stand_file, _csv_treatment_file, {}


if __name__ == '__main__':
    project_name = 'FHF24-0027-01 Kambo'
    project_settings_file = 'C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\ProjectForestsSettings.xlsx'

    # Set to False after Heureka simulation results have been saved in result_file, and you want to
    # rearrange the results so that they are easier to include in Excel calculations
    fix_import = False

    verbose = True

    project_folder, stand_file, average_over, stand_id_key, result_file, result_sheets, combine_sheets, \
        monetization_file, csv_stand_file, csv_treatment_file, kwargs = \
        project_settings(project_name, project_settings_file, fix_import)

    if fix_import:
        arrange_import(stand_file, csv_stand_file, csv_treatment_file, average_over, stand_id_key, project_name,
                       _verbose=verbose, **kwargs)
    elif not fix_import:
        arrange_results(result_file, result_sheets, combine_sheets, monetization_file, _verbose=verbose,  **kwargs)

