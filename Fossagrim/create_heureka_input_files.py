"""
Command line script that creates the input files to Heureka, and an empty Excel file for the Heureka results.

Usage:
    python create_heureka_input_files.py <project_name> <project_folder> <stand_file> [<unique_id>]

    project_name:    Name that identifies the current project, e.g. "FHF23-007 v01"
    project_folder:  Full path name to the folder where the inputs and outputs should be stored
    stand_file:      Full path file name of the "Forvaltningsplan" file (aka "Bestandsutvalg")
    unique_id:       Optional. Stored in the column "UniqueID" (column DA when opened in Excel) of the
                     averaged stand data. It is used as an identifier in the Woods platform to allow us to
                     separate between different versions

Creates, in project_folder:
    "<project_name> Averaged stand data.csv"   The stand register to import into Heureka. One area weighted
                                               average stand per tree species, over the active stands
    "<project_name> Heureka results.xlsx"      Empty Excel file with one sheet per tree species and method
                                               (BAU / PRES), where the Heureka results are pasted

After the simulations, run rearrange_heureka_results.py with the same arguments.
The script can be called from any folder.
"""
import argparse
import os
import sys

# Make the Fossagrim package importable when this file is run as a script
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import Fossagrim.io.fossagrim_io as fio  # noqa: E402
from Fossagrim.utils.definitions import DEFAULT_STAND_ID_KEY, FORVALTNINGSPLAN_DATA_SHEET  # noqa: E402
from Fossagrim.utils.project_setup import (  # noqa: E402
    averaged_stand_csv_file, heureka_results_file, build_result_sheets, build_combine_sheets,
    create_empty_results_file)


def project_settings(_project_name, _project_folder, _stand_file, create_output=True, _unique_id=None):
    """
    Works out the stands, file names and sheet names of a project, and optionally creates the input
    file for Heureka and the empty results file (see the module doc string).

    :param _project_name:
        str
        Name that identifies the current project, e.g. "FHF23-007 v01"
    :param _project_folder:
        str
        Full path name to the folder where the inputs and outputs should be stored
    :param _stand_file:
        str
        Full path file name of the "Forvaltningsplan" file (aka "Bestandsutvalg")
    :param create_output:
        bool
        When True the output files are created (existing files are overwritten)
    :param _unique_id:
        str or int
        When not None, it is stored in column "UniqueID" (column DA when opened in Excel) for each averaged stand.
        It is used as an identifier in the Woods platform to allow us to separate between different versions
    :return:
        tuple of
        _project_folder:  str
        _stand_file:      str
        _average_over:    dict, {'Spruce': [stand ids], 'Pine': [...]} of the active stands
        _stand_id_key:    str
        _result_file:     str, full path name of the Heureka results file
        _result_sheets:   list of str, sheet names in the results file
        _combine_sheets:  dict, see project_setup.build_combine_sheets()
        _csv_stand_file:  str, full path name of the stand register csv file
    """
    _stand_id_key = DEFAULT_STAND_ID_KEY
    _csv_stand_file = averaged_stand_csv_file(_project_folder, _project_name)
    _result_file = heureka_results_file(_project_folder, _project_name)

    _average_over, _active_prod_areas, _active_prod_areas_fract = fio.get_active_forests_from_stand(
        _stand_file, stand_id_key=_stand_id_key, sheet_name=FORVALTNINGSPLAN_DATA_SHEET)

    forest_types = list(_active_prod_areas_fract.keys())
    _result_sheets = build_result_sheets(_project_name, forest_types)

    for key, value in _active_prod_areas.items():
        print('{} stands have a total area of {} daa, which corresponds to a fraction of {:.3} of the total area'.format(
            key, sum(value), _active_prod_areas_fract[key]))

    _combine_sheets = build_combine_sheets(
        _project_name, forest_types, list(_active_prod_areas_fract.values()), _result_sheets)

    if create_output:
        create_empty_results_file(_result_file, _result_sheets)
        fio.export_fossagrim_stand_to_heureka(
            _stand_file,
            _csv_stand_file,
            average_over=_average_over,
            stand_id_key=_stand_id_key,
            average_name='{} '.format(_project_name),
            sheet_name=FORVALTNINGSPLAN_DATA_SHEET,
            verbose=False,
            unique_id=_unique_id
        )

    return _project_folder, _stand_file, _average_over, _stand_id_key, _result_file, _result_sheets, _combine_sheets, \
        _csv_stand_file


def main(argv=None):
    parser = argparse.ArgumentParser(description='Create the input files to Heureka, and an empty results file')
    parser.add_argument("project_name", help='Name that identifies the project, e.g. "FHF23-007 v01"')
    parser.add_argument("project_folder", help='Folder where the inputs and outputs are stored')
    parser.add_argument("stand_file", help='Full path name of the Forvaltningsplan Excel file')
    parser.add_argument("unique_id", nargs='?', help='Optional version identifier used by the Woods platform')
    args = parser.parse_args(argv)

    project_settings(args.project_name, args.project_folder, args.stand_file, True, _unique_id=args.unique_id)


if __name__ == '__main__':
    main()
