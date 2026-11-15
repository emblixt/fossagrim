"""
Naming conventions for the files and Excel sheets of a project, shared by create_heureka_input_files.py,
rearrange_heureka_results.py and utils/projects_new.py.

For a project named e.g. "FHF23-007 v01" with Spruce and Pine stands, the files in the project folder are:
    "FHF23-007 v01 Averaged stand data.csv"   stand register that is imported into Heureka
    "FHF23-007 v01 Heureka results.xlsx"      one sheet per simulation, where the Heureka results are pasted:
        "FHF23-007 v01 Spruce BAU", "FHF23-007 v01 Spruce PRES", "FHF23-007 v01 Pine BAU", "FHF23-007 v01 Pine PRES"
and the combined results (written by rearrange_raw_heureka_results) are named
    "FHF23-007 v01 Combined Stands Spruce-and-Pine BAU", "FHF23-007 v01 Combined Stands Spruce-and-Pine PRES"
"""
import os

from Fossagrim.utils.definitions import METHODS, EXCEL_MAX_SHEET_NAME_LENGTH

# Project names of this length, or longer, are truncated in the sheet names
MAX_PROJECT_NAME_LENGTH_IN_SHEETS = 20
TRUNCATED_PROJECT_NAME_LENGTH = 10


def averaged_stand_csv_file(project_folder, project_name):
    """Full path name of the stand register csv file that is imported into Heureka"""
    return os.path.join(project_folder, '{} Averaged stand data.csv'.format(project_name))


def heureka_results_file(project_folder, project_name):
    """Full path name of the Excel file where the Heureka results are stored"""
    return os.path.join(project_folder, '{} Heureka results.xlsx'.format(project_name))


def sheet_name_prefix(project_name, verbose=True):
    """
    The project name as used in sheet names. Excel can only handle sheet names that are at most 31 characters,
    and the longest string we add is " Spruce PRES" (12 characters), so long project names are truncated.
    """
    if len(project_name) >= MAX_PROJECT_NAME_LENGTH_IN_SHEETS:
        short_name = project_name[:TRUNCATED_PROJECT_NAME_LENGTH]
        if verbose:
            print('WARNING! Project name: "{}" is too long to fit in Excel sheet names, and is shortened to "{}"'.format(
                project_name, short_name))
        return short_name
    return project_name


def build_result_sheets(project_name, forest_types, verbose=True):
    """
    Names of the sheets in the Heureka results file, one per forest type and method, ordered as
        [<type 1> BAU, <type 1> PRES, <type 2> BAU, <type 2> PRES, ...]

    :param project_name:
        str
    :param forest_types:
        list of str, e.g. ['Spruce', 'Pine']
    :return:
        list of str
    """
    prefix = sheet_name_prefix(project_name, verbose=verbose)
    sheets = ['{} {} {}'.format(prefix, forest_type, method) for forest_type in forest_types for method in METHODS]
    for sheet in sheets:
        if len(sheet) > EXCEL_MAX_SHEET_NAME_LENGTH:
            raise ValueError('Sheet name "{}" is longer than the {} characters Excel allows'.format(
                sheet, EXCEL_MAX_SHEET_NAME_LENGTH))
    return sheets


def build_combine_sheets(project_name, forest_types, fractions, result_sheets):
    """
    Describes how the results of the different forest types are combined into one result per method, in the
    format used by heureka_results.rearrange_raw_heureka_results():
        {'<project_name> Combined Stands Spruce-and-Pine BAU': ['<Spruce BAU sheet>', 0.3, '<Pine BAU sheet>', 0.7],
         '<project_name> Combined Stands Spruce-and-Pine PRES': ['<Spruce PRES sheet>', 0.3, '<Pine PRES sheet>', 0.7]}

    :param project_name:
        str
    :param forest_types:
        list of str
    :param fractions:
        list of float, the fraction of the total active area of each forest type.
        When there is only one forest type, its fraction is set to 1
    :param result_sheets:
        list of str, as returned by build_result_sheets()
    :return:
        dict
    """
    if len(fractions) < 2:
        fractions = [1.]

    combine_sheets = {}
    for j, method in enumerate(METHODS):
        sheets_and_fractions = []
        for k, fraction in enumerate(fractions):
            sheets_and_fractions.append(result_sheets[len(METHODS) * k + j])
            sheets_and_fractions.append(fraction)
        combine_sheets['{} Combined Stands {} {}'.format(project_name, '-and-'.join(forest_types), method)] = \
            sheets_and_fractions
    return combine_sheets


def create_empty_results_file(result_file, result_sheets):
    """Creates (overwrites) an Excel file with the given, empty, sheets"""
    import pandas as pd

    writer = pd.ExcelWriter(result_file, engine='xlsxwriter')
    for sheet in result_sheets:
        writer.book.add_worksheet(sheet)
    writer.close()
