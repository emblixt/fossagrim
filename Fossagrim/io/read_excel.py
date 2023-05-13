import pandas as pd
import numpy as np


def load_rows(filename, header, sheet_name=None, year_key=None, data_key=None):
    """
    Loads a typical output from Heureka where a time serie is given as a row.
    At line header + 1, there must be cells containing 'Variable' and 'Period ...' that we use to
    identify the time line of our selected data

    :param filename:
    :param header:
    :param sheet_name:
    :param year_key:
    :param data_key:
    :return:
    """
    if sheet_name is None:
        sheet_name = 'Sheet1'
    if year_key is None:
        year_key = 'Year'
    if data_key is None:
        data_key = 'Total Carbon Stock (dead wood, soil, trees, stumps and roots)'

    table = pd.read_excel(filename, header=header, sheet_name=sheet_name, engine='openpyxl')
    table_keys = list(table.keys())
    variable_keys = list(table['Variable'])

    year_index = variable_keys.index(year_key)
    data_index = variable_keys.index(data_key)

    output_years = []
    output_data = []

    for key in table_keys:
        if 'Period' not in key:
            continue
        output_years.append(table[key][year_index])
        output_data.append(table[key][data_index])

    return np.array(output_years), np.array(output_data)

