"""
Reading and rearranging the results of Heureka simulations.

Raw Heureka results
-------------------
A simulation result is copied from Heureka (or exported from the Woods platform) into its own sheet of the
"<project> Heureka results.xlsx" file. Each sheet has one row per variable:

    | Variable                 | <category> | Unit  | Period 0 | Period 1 | ...
    | Year                     |            | year  | 2025     | 2030     | ...
    | Treatment                |            |       | None     | Thinning | ...
    | Soil Carbon Stock        |            | ton/ha| 50.1     | 50.3     | ...

Some treatments split a 5 year period into sub-periods; those columns are dropped when reading.

Rearranged results
------------------
rearrange_raw_heureka_results() transposes the raw sheets (one row per period, one column per variable),
writes them side by side into the sheet 'Rearranged results', and adds area-weighted combinations of the
forest types (e.g. Spruce and Pine) of each method (BAU / PRES).
"""
import os

import numpy as np
import pandas as pd

from Fossagrim.io.common import get_row_index, my_str, read_excel, write_csv_file
from Fossagrim.utils.definitions import (
    heureka_standdata_keys, heureka_standdata_desc, HEUREKA_PERIOD_LENGTH, REARRANGED_RESULTS_SHEET,
    HEUREKA_CSV_ENCODING)
from Fossagrim.utils.monetization_parameters import variables_used_in_monetization

TOTAL_CARBON = 'Total Carbon Stock (dead wood, soil, trees, stumps and roots)'
SOIL_CARBON = 'Soil Carbon Stock'
RECREATION_INDEX = 'Recreation Index After'
DEAD_STANDING_TREES = 'Dead Standing Trees >=20cm'
DOWNED_DEADWOOD = 'Downed Deadwood >=20cm'

# Columns that are copied (not area weighted) when forest types are combined
NON_NUMERIC_VARIABLES = ['Treatment', 'Year', 'UserDefined8 : UniqueID']

# Column with the first period in a raw Heureka results sheet; columns before it are Variable, category, Unit
FIRST_PERIOD_COLUMN = 3
UNIT_COLUMN = 2

DEFAULT_PROJECTS_DIR = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger"


def _qc_plot_dir(filename):
    qc_plot_dir = os.path.join(os.path.dirname(filename), 'QC_plots')
    os.makedirs(qc_plot_dir, exist_ok=True)
    return qc_plot_dir


def read_raw_heureka_results(filename, sheet_name, read_only_these_variables=None, verbose=False):
    """
    Read heureka results as exported from Heureka (by simply copying a simulation result and pasting it
    into an empty sheet in Excel) and returns the data in a transposed DataFrame with one row for each
    5 year period. Sub-periods (years that are not a multiple of 5) are dropped.
    NOTE, the variable 'Year' must be included in the raw results

    :param filename:
        str
        Full path name of the Heureka results Excel file
    :param sheet_name:
        str
    :param read_only_these_variables:
       list
       If provided, only the variables (given by their exact name) contained in this list are read
    :param verbose:
        bool
        creates qc plot(s) in the folder QC_plots next to 'filename'
    :return:
        pandas DataFrame with one column per variable, or None if the sheet could not be read.
        The units of the variables are stored in the DataFrame.attrs dictionary
    """
    table = read_excel(filename, 0, sheet_name)
    if table is None:
        return None
    if 'Variable' not in table.keys():
        print('Column "Variable" not found in sheet {} of {}'.format(sheet_name, os.path.basename(filename)))
        return None

    year_row = get_row_index(table, 'Variable', 'Year', exact=True)
    if year_row is None:
        year_row = get_row_index(table, 'Variable', 'Year')
    if year_row is None:
        print('Variable Year not in sheet {} of {}'.format(sheet_name, os.path.basename(filename)))
        return None

    # Some treatments in Heureka break the default periods of 5 years into smaller sub-periods, we choose to
    # remove these from the returned dataframe
    five_years = [np.mod(float(this_year), HEUREKA_PERIOD_LENGTH) == 0
                  for this_year in table.iloc[year_row, FIRST_PERIOD_COLUMN:]]

    data_dict = {}
    unit_dict = {}
    for row_i, variable in enumerate(table['Variable']):
        if not isinstance(variable, str):
            continue  # empty row
        if read_only_these_variables is not None and variable not in read_only_these_variables:
            continue
        if variable in data_dict:
            continue  # duplicated variable name, keep the first one
        data_dict[variable] = table.iloc[row_i, FIRST_PERIOD_COLUMN:][five_years]
        unit_dict[variable] = table.iloc[row_i, UNIT_COLUMN]

    if verbose:
        from Fossagrim.plotting.misc_plots import plot_raw_data
        plot_raw_data(data_dict, sheet_name, 'Raw Data', _qc_plot_dir(filename))

    result = pd.DataFrame(data=data_dict)
    result.attrs = unit_dict
    return result


def combine_raw_heureka_results(filename, sheet_name, function, variables, verbose=False):
    """
    Uses read_raw_heureka_results() to load all variables from a raw Heureka result,
    then combines the variables according to the input function, eg:
    > variables = {'x': 'Total Carbon Stock (dead wood, soil, trees, stumps and roots)',
                   'y': 'Soil Carbon Stock'}
    > def func(x, y):
        return x - y
    > combine_raw_heureka_results(filename, sheet_name, func, variables)

    :param filename:
    :param sheet_name:
    :param function:
        function whose parameter names are the keys of 'variables'
    :param variables
        dict
        Dictionary of variable where each key is a parameter in the input function (e.g. 'x') and the
        value is the name of variable in the raw heureka results (e.g. 'Soil Carbon Stock')
    :param verbose:
    :return:
        the output of 'function'
    """
    data = read_raw_heureka_results(filename, sheet_name, verbose=verbose)
    _variables = {_key: data[variables[_key]] for _key in variables}
    return function(**_variables)


def combine_tables(tables, fractions):
    """
    Area weighted combination of several rearranged Heureka results (DataFrames with the same columns).
    Columns in NON_NUMERIC_VARIABLES are copied from the first table.

    :param tables:
        list of pandas DataFrames
    :param fractions:
        list of floats, one per table (normally summing to 1)
    :return:
        dict of column name: values
    """
    combined_dict = {}
    for key in tables[0].keys():
        if key in NON_NUMERIC_VARIABLES:
            combined_dict[key] = tables[0][key].values
            continue
        this_column = np.zeros(len(tables[0][key]))
        for table, fraction in zip(tables, fractions):
            this_column = this_column + np.array([float(_x) for _x in table[key].values]) * fraction
        combined_dict[key] = this_column
    return combined_dict


def _write_table(filename, table, sheet_name, start_col, title):
    """Write 'table' to 'sheet_name' of an existing Excel file, starting at row 3, with 'title' in row 1"""
    from openpyxl import load_workbook

    with pd.ExcelWriter(filename, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
        table.to_excel(writer, sheet_name=sheet_name, startcol=start_col, startrow=2)
    wb = load_workbook(filename)
    wb[sheet_name].cell(1, start_col + 1).value = title
    wb.save(filename)


def rearrange_raw_heureka_results(filename, sheet_names, combine_sheets, monetization_file=None, verbose=False):
    """
    Transposes the raw Heureka results in 'sheet_names' and writes them side by side in the sheet
    'Rearranged results' of the same file, followed by the area weighted combinations given in 'combine_sheets'.

    Each table occupies (number of variables + 1) columns, followed by one empty column. Its name is written in
    row 1, the variable names in row 3 and the data from row 4.

    If the 'Rearranged results' sheet already exists, nothing is written to 'filename' (remove the sheet and
    try again).

    :param filename:
        str
        Name of file with Heureka results
    :param sheet_names:
        list
        List of strings with names of sheets that contains raw results from different Heureka simulations
    :param combine_sheets:
        dict
        { <NAME OF COMBINED RESULTS>: [<SHEET NAME FOR FOREST TYPE 1 IN HEUREKA RAW RESULTS>, <FRACTION TYPE 1>,
                                       <SHEET NAME FOR FOREST TYPE 2 IN HEUREKA RAW RESULTS>, <FRACTION TYPE 2>,
                                       ... ],
         <ANOTHER COMBINED RESULT>: [...]}
        See utils/project_setup.py build_combine_sheets()
    :param monetization_file:
        str
        Full path name of monetization file which we try to generate. The user is asked before an existing
        file is overwritten
    :param verbose:
        bool
        creates qc plot(s)
    :return:
        bool
        True if the monetization file was written
    """
    from openpyxl import load_workbook

    wb = load_workbook(filename)
    rearrange_result_file = REARRANGED_RESULTS_SHEET not in wb.sheetnames
    wb.close()
    if not rearrange_result_file:
        # The data in an old sheet is not removed before overwriting, so old data could hang on also after
        # 'over writing'. Therefore, we refuse to overwrite.
        print("WARNING the sheet '{}' already exists in {}. Remove it before trying again.".format(
            REARRANGED_RESULTS_SHEET, os.path.basename(filename)))

    write_monetization_file = False
    if monetization_file is not None:
        if os.path.isfile(monetization_file):
            print("WARNING monetization file {} already exists.".format(os.path.basename(monetization_file)))
            _ans = input("Do you want to overwrite? Y/[N]:")
            write_monetization_file = _ans.upper() == 'Y'
        else:
            write_monetization_file = True

        if write_monetization_file:
            # Create empty monetization file
            writer = pd.ExcelWriter(monetization_file, engine='xlsxwriter')
            for _sheet in [REARRANGED_RESULTS_SHEET, 'Parameters', 'Monetization']:
                writer.book.add_worksheet(_sheet)
            writer.close()

    start_cols = []
    for i, sheet_name in enumerate(sheet_names):
        table = read_raw_heureka_results(
            filename, sheet_name, read_only_these_variables=variables_used_in_monetization, verbose=verbose)
        if table is None:
            raise ValueError('Could not read Heureka results from sheet {} in {}'.format(
                sheet_name, os.path.basename(filename)))
        this_start_col = i * (len(table.keys()) + 2)
        if rearrange_result_file:
            _write_table(filename, table, REARRANGED_RESULTS_SHEET, this_start_col, sheet_name)
        start_cols.append(this_start_col)

    if combine_sheets is None:
        return write_monetization_file

    last_start_col = start_cols[-1]
    for i, (this_combined_result, sheets_and_fractions) in enumerate(combine_sheets.items()):
        # sheet names are found in positions 0, 2, 4, ... and the fractions in positions 1, 3, 5, ...
        combine_sheet_names = sheets_and_fractions[0::2]
        fractions = sheets_and_fractions[1::2]
        tables = [read_raw_heureka_results(filename, sheet_name, variables_used_in_monetization)
                  for sheet_name in combine_sheet_names]
        combined_dict = combine_tables(tables, fractions)
        combined_table = pd.DataFrame(data=combined_dict)

        n_cols = len(combined_table.keys()) + 2
        if rearrange_result_file:
            _write_table(filename, combined_table, REARRANGED_RESULTS_SHEET,
                         last_start_col + (i + 1) * n_cols, this_combined_result)

        if write_monetization_file:
            _write_table(monetization_file, combined_table, REARRANGED_RESULTS_SHEET,
                         i * n_cols, this_combined_result)

        if verbose:
            from Fossagrim.plotting.misc_plots import plot_raw_data
            plot_raw_data(combined_dict, this_combined_result, 'Combined', _qc_plot_dir(filename))

    return write_monetization_file


def read_rearranged_heureka_results(filename, sheet_name=None):
    """
    Reads the sheet written by rearrange_raw_heureka_results() and returns a dictionary which has each model
    name (e.g. 'FHF24-0047 v02 Spruce BAU') as key and the associated data as value. The data is a pandas
    DataFrame which has the same columns as the output from read_raw_heureka_results()

    :param filename:
    :param sheet_name:
        Default 'Rearranged results'
    :return:
        dict
    """
    if sheet_name is None:
        sheet_name = REARRANGED_RESULTS_SHEET

    table = read_excel(filename, 0, sheet_name)
    # The model names are the only non-empty cells in the first row
    models = {}
    for _i, _x in enumerate(table.keys()):
        if isinstance(_x, str) and 'Unnamed' in _x:
            continue
        models[_x] = _i
    model_names = list(models.keys())
    first_col = models[model_names[0]]
    if len(model_names) > 1:
        # distance between two model names = index column + variables + one empty column
        block_width = models[model_names[1]] - first_col
    else:
        block_width = len(table.keys()) - first_col + 1
    n_params = block_width - 1  # index column + variables

    result = {}
    for _i, model in enumerate(model_names):
        _tmp = {}
        start_c = first_col + _i * block_width
        end_c = start_c + n_params
        for _c in range(start_c + 1, end_c):  # + 1 because we skip the index column
            param = table.iloc[1, _c]
            _tmp[param] = table.iloc[2:, _c]
        result[model] = pd.DataFrame(data=_tmp)
    return result


def get_nature_and_climate_effect(result_file, stand_id, average_years=None, verbose=False):
    """
    Calculates the difference in predefined climate and nature parameters, between PRES and BAU,
    and sums the result over the given number of years.

    The sheets '<stand_id> PRES' and '<stand_id> BAU' must exist in 'result_file'.

    :param result_file:
         str
         Name of excel file that contain the "raw" Heureka results
    :param stand_id
         str
    :param average_years:
         int
         Number of years to sum the effect over. Default 30
    :param verbose:
         bool
    :return:
        tuple (carbon, dead wood, recreation)
        carbon:     sum of the difference in total carbon, ignoring soil carbon
        dead wood:  sum of the difference in standing + downed dead wood >= 20 cm, or None if not in the results
        recreation: sum of the difference in recreation index, or None if not in the results
        (None, None, None) if any of the two sheets could not be read
    """
    if average_years is None:
        average_years = 30

    i = int(np.floor(average_years / HEUREKA_PERIOD_LENGTH))
    if verbose:
        print('average_years: {} corresponds to {} periods'.format(average_years, i))

    results = {}
    for method in ['PRES', 'BAU']:
        try:
            results[method] = read_raw_heureka_results(result_file, '{} {}'.format(stand_id, method), verbose=verbose)
        except ValueError as error:  # e.g. missing sheet
            print(error)
            results[method] = None
        if results[method] is None:
            return None, None, None
    pres = results['PRES']
    bau = results['BAU']

    def values(table, key):
        return np.array([float(_x) for _x in table[key].values])

    _c_diff = values(pres, TOTAL_CARBON) - values(pres, SOIL_CARBON) - values(bau, TOTAL_CARBON) + \
        values(bau, SOIL_CARBON)
    c_diff = np.sum(_c_diff[:i])
    if verbose:
        print('Variables to sum over:')
        print('  Carbon (neglecting soil):', _c_diff[:i])

    r_diff = None
    if RECREATION_INDEX in pres.keys():
        _r_diff = values(pres, RECREATION_INDEX) - values(bau, RECREATION_INDEX)
        r_diff = np.sum(_r_diff[:i])
        if verbose:
            print('  Recreation index:', _r_diff[:i])

    d_diff = None
    if DEAD_STANDING_TREES in pres.keys():
        _d_diff = values(pres, DEAD_STANDING_TREES) + values(pres, DOWNED_DEADWOOD) - \
            values(bau, DEAD_STANDING_TREES) - values(bau, DOWNED_DEADWOOD)
        d_diff = np.sum(_d_diff[:i])
        if verbose:
            print('  Dead wood >= 20cm:', _d_diff[:i])

    return c_diff, d_diff, r_diff


def collect_all_stand_data(out_file=None, base_dir=None, dry_run=True):
    """
    Traverses the folders below 'base_dir' for "*Averaged stand data.csv" files that have a matching
    "*Heureka results.xlsx" file, calculates the nature and climate effects (get_nature_and_climate_effect())
    of each stand, and collects all stands, with these effects added, in one Heureka csv file.

    :param out_file:
        str
        Full file name of output file. Required when dry_run is False
    :param base_dir:
        Path to directory where Heureka modelling results are stored (in sub folders of base_dir)
    :param dry_run:
        bool
        If True, no results are written, just checking for potential Heureka files and calculating effects
    :return:
        None
    """
    from pathlib import Path

    if base_dir is None:
        base_dir = DEFAULT_PROJECTS_DIR
    if not dry_run and out_file is None:
        raise IOError('Name of output file must be given')

    effect_columns = [heureka_standdata_keys.index(_key) for _key in [
        'UserDefinedVariable5_ClimateEffect', 'UserDefinedVariable6_DeadWoodEffect',
        'UserDefinedVariable7_RecreationEffect']]

    # Creates a new Heureka compatible csv file for storing the data
    if not dry_run:
        write_csv_file(out_file, heureka_standdata_keys, heureka_standdata_desc, append=False)

    output_lines = []
    for stand_file in Path(base_dir).rglob('*Averaged stand data.csv'):
        print('Working on {}'.format(stand_file.name))
        heureka_result_file = str(stand_file).replace('Averaged stand data.csv', 'Heureka results.xlsx')
        if not os.path.isfile(heureka_result_file):
            continue
        print(' Heureka result file exists')

        with open(stand_file, 'r', encoding=HEUREKA_CSV_ENCODING) as _in:
            for line in _in.readlines():
                split_line = line.split(';')
                if 'FHF' not in split_line[0]:
                    continue
                print('  Searching for stand id: {}'.format(split_line[0]))
                effects = get_nature_and_climate_effect(heureka_result_file, split_line[0])
                if effects[0] is None:
                    alt_stand_id = split_line[0].replace('Avg Stand-', '')
                    print('  Searching for alternative stand id: {}'.format(alt_stand_id))
                    effects = get_nature_and_climate_effect(heureka_result_file, alt_stand_id)
                print('  ', effects[0])
                for column, effect in zip(effect_columns, effects):
                    split_line[column] = my_str(effect)
                output_lines.append(';'.join(split_line))

    if dry_run:
        print('Dry run: {} stands found, nothing written'.format(len(output_lines)))
        return

    with open(out_file, 'a', encoding=HEUREKA_CSV_ENCODING) as _out:
        _out.writelines(output_lines)
