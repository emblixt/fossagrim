"""
Script than can be called from command line to create the necessary input files to Heureka, and an empty
results file
"""
import argparse
import sys


def project_settings(_project_name, _project_folder, _stand_file):
    """
    Creates the necessary input files and results file to make it easy to interact with Heureka using our
    internal "Forvaltningsplan" file set up.

    Keys expected in project settings files:
    "Project name", "Project folder", "Status", "Stand id key", "Stands file"

    :param _project_name:
        str
        Name that identifies the current project, e.g. "FHF23-007 v01"

    :param _project_folder:
        str
        Full path name to the folder where the inputs and outputs should be stored

    :param _stand_file:
        str
        Full path file name of the "Forvaltningsplan" file (aka "Bestandsutvalg")

    :return:
        None
        Creates two files:
            "<_project_name> Averaged stand data.csv"  which is the input file to Heureka
            "<_project_name> Heureka results.xlsx"  which is an empty xlsx file (with specific sheet names) that we
            will use to store results from Heureka in later


    """
    import os
    import numpy as np
    import pandas as pd
    sys.path.append('..')

    import Fossagrim.io.fossagrim_io as fio

    print(os.path.dirname(__name__))

    methods = ['BAU', 'PRES']  # "Business as usual" and "Preservation"

    _stand_id_key = 'Fossagrim ID'
    _csv_stand_file = os.path.join(_project_folder, '{} Averaged stand data.csv'.format(_project_name))
    # Full path file name of the empty Heureka results file that we shall save the Heureka results in
    _result_file = os.path.join(_project_folder, '{} Heureka results.xlsx'.format(_project_name))

    _average_over, _active_prod_areas, _active_prod_areas_fract = fio.get_active_forests_from_stand(
        _stand_file, stand_id_key=_stand_id_key, sheet_name='data')

    # Excel can only handle sheet names that are shorter than 31 characters, and the longest string
    # we add is "Spruce PRES", which is 11 characters long.
    if len(_project_name) >= 20:
        print('WARNING! Project name: {} is too long and has to be shortened to fit in Excel')
        tmp_project_name = _project_name[:10]
    else:
        tmp_project_name = _project_name
    _result_sheets = \
        ['{} {} {}'.format(tmp_project_name, _x, _y) for _x, _y in
         zip(np.repeat(list(_average_over.keys()), 2), methods * len(_average_over))]

    for key, value in _active_prod_areas.items():
        print('{} stands have a total area of {} daa, which corresponds to a fraction of {:.3} of the total area'.format(
            key, sum(value), _active_prod_areas_fract[key]))

    _combine_sheets = {}

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

    # Create empty results file
    writer = pd.ExcelWriter(_result_file, engine='xlsxwriter')
    wb = writer.book
    for _sheet in _result_sheets:
        _ = wb.add_worksheet(_sheet)
    writer.close()

    table = fio.read_excel(_stand_file, 7, 'data')

    fio.export_fossagrim_stand_to_heureka(
        _stand_file,
        _csv_stand_file,
        average_over=_average_over,
        stand_id_key=_stand_id_key,
        average_name='{} '.format(args.project_name),
        sheet_name='data',
        verbose=False
    )

    return _project_folder, _stand_file, _average_over, _stand_id_key, _result_file, _result_sheets, _combine_sheets, \
        _csv_stand_file


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    """
    :param project_name:
    str
    Name that identifies the current project, e.g. "FHF23-007 v01"

    :param project_folder:
    str
    Full path name to the folder where the inputs and outputs should be stored

    :param stand_file:
    str
    Full path file name of the "Forvaltningsplan" file (aka "Bestandsutvalg")
    """
    parser.add_argument("project_name")
    parser.add_argument("project_folder")
    parser.add_argument("stand_file")
    args = parser.parse_args()

    project_folder, stand_file, average_over, stand_id_key, result_file, result_sheets, combine_sheets, \
        csv_stand_file = \
        project_settings(args.project_name, args.project_folder, args.stand_file)

