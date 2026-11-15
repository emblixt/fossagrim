"""
Low level helpers for reading Excel files and reading/writing the semicolon separated csv files used by Heureka.

Heureka csv layout (e.g. "StandData.csv", "<project> Averaged stand data.csv"):
    line 1: a description of each column (may be empty strings)
    line 2: the column keys, e.g. heureka_standdata_keys in utils/definitions.py
    line 3...: one line of data per stand
"""
import numpy as np
import pandas as pd

from Fossagrim.utils.definitions import HEUREKA_CSV_ENCODING


def my_float(x):
    """
    float() that also accepts strings with a decimal comma, e.g. '1,5' -> 1.5.
    Excel files edited with Norwegian locale often contain a mixture of '.' and ',' as decimal sign.
    """
    if isinstance(x, str):
        return float(x.replace(',', '.'))
    return float(x)


def my_str(x):
    """str() that returns an empty string for None."""
    if x is None:
        return ''
    return str(x)


def my_str_to_float(x):
    """
    Convert a string to int (if it represents a whole number) or float. Strings that are not numbers are
    returned unchanged, e.g. '3' -> 3, '3.5' -> 3.5, 'abc' -> 'abc', '' -> ''
    """
    try:
        a = float(x)
        if a.is_integer():
            a = int(a)
    except (ValueError, TypeError):
        a = x
    return a


def get_row_index(table, key, item, exact=False):
    """
    Return the (zero-based) row index of the first row where column 'key' contains 'item'.

    :param table:
        pandas DataFrame
    :param key:
        str
        Name of the column to search within
    :param item:
        str
        The text to look for
    :param exact:
        bool
        If True the cell must be equal to 'item', otherwise it is enough that 'item' is a substring of the cell
    :return:
        int or None if no match. Non-string cells (e.g. empty cells read as NaN) are ignored.
    """
    for i, var in enumerate(table[key]):
        if not isinstance(var, str):
            continue
        if (exact and var.strip() == item) or (not exact and item in var):
            return i
    return None


def read_excel(filename, header, sheet_name):
    """
    Read one sheet of an Excel file into a pandas DataFrame, and strip leading/trailing white space (including
    line breaks) from the column names.

    :param filename:
        str
    :param header:
        int
        Zero-based row index of the row that holds the column names
    :param sheet_name:
        str or int
    :return:
        pandas DataFrame, or None if the file could not be opened because it is locked (typically because
        it is open in Excel)
    """
    try:
        table = pd.read_excel(filename, header=header, sheet_name=sheet_name, engine='openpyxl')
    except PermissionError as err_msg:
        print(err_msg)
        print('Is the file open in Excel? Close it and try again.')
        return None

    table = table.rename(columns={_key: _key.strip() for _key in table.keys() if isinstance(_key, str)})
    return table


def write_csv_file(write_to_file, default_keys, default_desc, append=False, **kwargs):
    """
    Writes the parameter values from the kwargs as one line in a Heureka csv file (see module doc string).

    Example output is the StandData.csv files that is used by Heureka for import of stand data

    :param write_to_file:
        Name of csv file to write data to
    :param default_keys:
        list
        List of key names that the csv file must include, eg. heureka_standdata_keys
    :param default_desc:
        list
        List of descriptions that the csv file can include, eg. heureka_standdata_desc
    :param append:
        bool
        if True, the data in kwargs are appended to an existing output file without creating a header.
        if False, any existing file is overwritten and the two header lines are written
    :param kwargs:
        keywords must be among 'default_keys'. Unknown keys are reported and ignored.
        Floats are written with two decimals, everything else with str().
        If no kwargs are given, only the header lines are written (when append is False)
    :return:
        None
    """
    if len(default_keys) != len(default_desc):
        raise IOError('The lists of keys and descriptions must be of same length')

    if not append:
        with open(write_to_file, 'w', encoding=HEUREKA_CSV_ENCODING) as f:
            f.write(';'.join(default_desc) + '\n')
            f.write(';'.join(default_keys) + '\n')

    data = [''] * len(default_keys)
    for key, this_data in kwargs.items():
        if key not in default_keys:
            print('WARNING: key "{}" not found in accepted default keys'.format(key))
            continue
        if isinstance(this_data, (float, np.floating)):
            this_string = '{:.2f}'.format(this_data)
        else:
            this_string = str(this_data)
        data[default_keys.index(key)] = this_string

    if len(kwargs) > 0:
        with open(write_to_file, 'a', encoding=HEUREKA_CSV_ENCODING) as f:
            f.write(';'.join(data) + '\n')


def read_csv_file(read_from_file):
    """
    Reads a Heureka csv file (see module doc string), e.g. a StandData.csv file that is used as input to Heureka

    :param read_from_file:
        str
    :return:
        dict, list
        A dictionary with each column key as dictionary key, and the values of that column (one per stand,
        converted to int/float where possible) as a list.
        And the list of column descriptions, which makes it easy to use the output to write a new csv file
    """
    out = {}
    descriptors = []
    with open(read_from_file, 'r', encoding=HEUREKA_CSV_ENCODING) as _in:
        for i, line in enumerate(_in.readlines()):
            split_line = line.rstrip('\r\n').split(';')
            if i == 0:
                descriptors = split_line
            elif i == 1:
                out = {_key: [] for _key in split_line}
            else:
                keys = list(out.keys())
                for j, val in enumerate(split_line):
                    out[keys[j]].append(my_str_to_float(val))

    return out, descriptors
