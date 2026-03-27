"""
Script than can be called from command line to rearrange the heureka results
"""
from create_heureka_input_files import project_settings
import argparse
import sys

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
    
    :param heureka_results_file:
    str
    Full path file name excel file that contains the heureka results
    
    Use it by calling:
    > python rearrange_heureka_results.py <project_name> <project_folder> <stand_file> <heureka_results_file>
    NOTE!
    You must be located in the same folder as the rearrange_heureka_results.py script for the relative import to work
    """
    sys.path.append('..')
    import Fossagrim.io.fossagrim_io as fio

    parser.add_argument("project_name")
    parser.add_argument("project_folder")
    parser.add_argument("stand_file")
    parser.add_argument("heureka_results_file")
    args = parser.parse_args()

    project_folder, stand_file, average_over, stand_id_key, _, result_sheets, combine_sheets, \
        csv_stand_file = \
        project_settings(args.project_name, args.project_folder, args.stand_file, False)

    fio.rearrange_raw_heureka_results(args.heureka_results_file, result_sheets, combine_sheets)