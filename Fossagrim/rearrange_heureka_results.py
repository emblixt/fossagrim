"""
Command line script that rearranges the Heureka results, after they have been pasted into the
"<project_name> Heureka results.xlsx" file created by create_heureka_input_files.py.

Usage:
    python rearrange_heureka_results.py <project_name> <project_folder> <stand_file> <heureka_results_file>

    project_name:          Name that identifies the current project, e.g. "FHF23-007 v01"
    project_folder:        Full path name to the folder where the inputs and outputs are stored
    stand_file:            Full path file name of the "Forvaltningsplan" file (aka "Bestandsutvalg")
    heureka_results_file:  Full path file name of the Excel file that contains the Heureka results

The first three arguments must be the same as used with create_heureka_input_files.py, so that the same sheet
names are found. The results are written to the sheet 'Rearranged results' of heureka_results_file: one table per
simulation, plus the area weighted combination of the tree species for each method (BAU / PRES).
The script can be called from any folder.
"""
import argparse
import os
import sys

# Make the Fossagrim package importable when this file is run as a script
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import Fossagrim.io.fossagrim_io as fio  # noqa: E402
from Fossagrim.create_heureka_input_files import project_settings  # noqa: E402


def main(argv=None):
    parser = argparse.ArgumentParser(description='Rearrange the Heureka results of a project')
    parser.add_argument("project_name", help='Name that identifies the project, e.g. "FHF23-007 v01"')
    parser.add_argument("project_folder", help='Folder where the inputs and outputs are stored')
    parser.add_argument("stand_file", help='Full path name of the Forvaltningsplan Excel file')
    parser.add_argument("heureka_results_file", help='Full path name of the Excel file with the Heureka results')
    args = parser.parse_args(argv)

    _, _, _, _, _, result_sheets, combine_sheets, _ = \
        project_settings(args.project_name, args.project_folder, args.stand_file, create_output=False)

    fio.rearrange_raw_heureka_results(args.heureka_results_file, result_sheets, combine_sheets)


if __name__ == '__main__':
    main()
