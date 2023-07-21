import pandas as pd
import numpy as np
import os

from Fossagrim.utils.definitions import heureka_mandatory_standdata_keys, heureka_standdata_keys, \
    heureka_standdata_desc, fossagrim_standdata_keys, translate_keys_from_fossagrim_to_heureka


def read_excel(filename, header, sheet_name):
    try:
        table = pd.read_excel(filename, header=header, sheet_name=sheet_name, engine='openpyxl')
    except PermissionError as err_msg:
        print(err_msg)
        return None
    return table


def load_heureka_results(filename, header, sheet_name=None, year_key=None, data_key=None):
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

    table = read_excel(filename, header, sheet_name)
    if table is None:
        return None
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


def arrange_export_of_one_stand(row_index, table):
    """
    Utility function that collects, translates, and converts, the Fossagrim data into a format
    that can be used by heureka
    :param row_index:
        int
        row index (beginning at 0) of table which we want to export
    :param table:
        pandas DataFrame
        Table with data as read in 'export_fossagrim_stand_to_heureka'
        Each row contains one stand
    :return:
        dict
        Dictionary of keyword: value pairs that can be used by 'write_heureka_standdata'
    """
    key_translation = translate_keys_from_fossagrim_to_heureka()
    keyword_arguments = {}

    # loop over all mandatory keywords that heureka needs as input
    for key in heureka_mandatory_standdata_keys:
        keyword_arguments[key] = ''  # Give every mandatory key a default empty value at first
        if key in list(key_translation.keys()):
            # Use the translated keyword to extract data from the Fossagrim table. NO CONVERSION YET!
            keyword_arguments[key] = table[key_translation[key]][row_index]

    # Now do the conversion from Fossagrim data and units to Heureka compatible data and units
    for key, value in list(keyword_arguments.items()):
        if key in ['InventoryYear', 'N', 'CountyCode', 'SoilMoistureCode', 'VegetationType', 'Peat']:
            keyword_arguments[key] = int(value)
        if key == 'ProdArea':
            keyword_arguments[key] = value * 0.1  # from daa to ha
        if key == 'SiteIndexSpecies':
            if value == 'Gran':
                keyword_arguments[key] = 'G'
            elif value == 'Furu':
                keyword_arguments[key] = 'T'
            else:
                raise NotImplementedError('Species "{}" is not recognized as either Gran or Furu'.format(value))
        if key == 'V':
            keyword_arguments[key] = value * 10.  # m3/daa to m3/ha
        if key == 'PropPine':
            keyword_arguments[key] = table['Furu'][row_index] / table['Total'][row_index]
        if key == 'PropSpruce':
            keyword_arguments[key] = table['Gran'][row_index] / table['Total'][row_index]
        if key == 'PropBirch':
            keyword_arguments[key] = table['Lauv'][row_index] / table['Total'][row_index]
        if key == 'StandId':
            keyword_arguments[key] = table['Fossagrim ID'][row_index]

    return keyword_arguments


def average_over_stands(average_over, table, stand_id_key, average_name):
    """
    Calculates the area weighted averages

    :param average_over:
        list
        List of strings which determines which stands to use in calculating an average stand.
    :param table:
        pandas DataFrame
        Table with data as read in 'export_fossagrim_stand_to_heureka'
        Each row contains one stand
    :param stand_id_key:
        str
        Key that is used to identify the stands used in 'average_over'
    :param average_name:
        str
        Given name of the calculated average stand
    :return:
        panda DataFrame
        A copy of the original input table, but with only one line of data, which is the
        area weighted averages over the selected stands
    """
    from Fossagrim.utils.definitions import fossagrim_keys_to_average_over as keys_to_average_over
    from Fossagrim.utils.definitions import fossagrim_keys_to_sum_over as keys_to_sum_over
    from Fossagrim.utils.definitions import fossagrim_keys_to_most_of as keys_to_most_of

    # create a new empty DataFrame that should hold the average values
    avg_table = pd.DataFrame(columns=list(table.keys()))

    # extract the row indexes over which the average is calculated
    average_ind = []
    for i, stand_id in enumerate(table[stand_id_key]):
        if stand_id in average_over:
            average_ind.append(i)

    # continue using the above indexes to calculate the average values
    total_area = np.sum(table['Prod.areal'][average_ind])
    for key in fossagrim_standdata_keys:
        if key == 'Fossagrim ID':
            avg_table[key] = [average_name]
        elif key in keys_to_sum_over:
            avg_table[key] = np.sum(table[key][average_ind])
        elif key in keys_to_average_over:
            avg_table[key] = np.sum(table['Prod.areal'][average_ind] * table[key][average_ind]) / total_area
        elif key in keys_to_most_of:
            # An area weighted 'most of' calculation
            this_data = table[key][average_ind]
            area_weighted_average = np.sum(this_data * table['Prod.areal'][average_ind]) / total_area
            # index of original data closest to the average
            ind = np.argmin((this_data - area_weighted_average)**2)
            avg_table[key] = this_data[ind]
        else:
            # for other parameters, just copy the first selected stand
            avg_table[key] = [table[key][average_ind[0]]]

    return avg_table


def export_fossagrim_stand_to_heureka(read_from_file, write_to_file, this_stand_only=None, average_over=None,
                                      stand_id_key=None,
                                      header=None, sheet_name=None):
    """
    Load a 'typical' Fossagrim stand data file (Bestandsutvalg) and writes an output file which Heureka can use

    :param read_from_file:
    :param write_to_file:
    :param this_stand_only:
        str
        Name of stand to export
    :param average_over:
        list
        List of strings which determines which stands to use in calculating an average stand.
    :param stand_id_key:
        str
        Key that is used to identify the stands used in 'average_over'
    :param header:
    :param sheet_name:
    :return:
    """
    if stand_id_key is None:
        stand_id_key = 'Bestand'
    if sheet_name is None:
        sheet_name = 0   # Reads only first sheet, opposite to pandas default were None reads all sheets.
    if header is None:
        header = 7

    table = read_excel(read_from_file, header, sheet_name)
    if table is None:
        return None

    append = False

    if (average_over is not None) and isinstance(average_over, list):  # export averaged stands
        average_table = average_over_stands(average_over, table, stand_id_key, 'Avg stand')
        # get the data from the fossagrim table arranged for writing in heureka format
        keyword_arguments = arrange_export_of_one_stand(0, average_table)
        write_heureka_standdata(
            write_to_file,
            append=append,
            Note='Area weighted average of stand {}'.format(', '.join([str(_x) for _x in average_over])),
            **keyword_arguments
        )

    else:  # Only export individual stands
        for i, stand_id in enumerate(table[stand_id_key]):
            if (table['Bonitering\ntreslag'][i] == 'Uproduktiv') or np.isnan(table['HovedNr'][i]):
                continue
            if (this_stand_only is not None) and (this_stand_only != stand_id):
                continue
            # get the data from the fossagrim table arranged for writing in heureka format
            keyword_arguments = arrange_export_of_one_stand(i, table)
            write_heureka_standdata(
                write_to_file,
                append=append,
                Note='test',
                **keyword_arguments
            )
            append = True


def write_heureka_standdata(write_to_file, append=False, **kwargs):
    """
    Writes the parameter values from the kwargs to the 'write_to_file'
    :param write_to_file:
        Name of csv file to write data to
    :param append:
        bool
        if True, the data in kwargs are appended to an existing output file without creating a header
    :param kwargs:
        list of keyword arguments.
        keywords must match the list 'heureka_standdata_keys' given in utils.definitions.py
        The value of each key is used when writing the 'write_to_file' file
    :return:
    """

    # overwrite any existing file, and create the header.
    if not append:
        with open(write_to_file, 'w') as f:
            f.write(';'.join(heureka_standdata_desc)+'\n')
            f.write(';'.join(heureka_standdata_keys)+'\n')

    # arrange the data from kwargs
    data = ['']*len(heureka_standdata_keys)
    for key in kwargs:
        if key not in heureka_standdata_keys:
            print('WARNING: key "{}" not found in accepted keys of Heureka stand data files'.format(key))
            continue
        data[heureka_standdata_keys.index(key)] = str(kwargs[key])

    # append the data
    with open(write_to_file, 'a') as f:
        f.write(';'.join(data)+'\n')


def test_write_heureka_standdata():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    write_heureka_standdata(os.path.join(dir_path, 'test.csv'), test=8, DGV=46, SiteIndexSpecies='G')
    write_heureka_standdata(os.path.join(dir_path, 'test.csv'), append=True, SiteIndexSpecies='T')


def test_export_fossagrim_stand():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    read_from_file = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-003 Kloppmyra\\FHF23-003 Bestandsutvalg.xlsx"
    write_to_file = os.path.join(dir_path, 'test.csv')

    export_fossagrim_stand_to_heureka(read_from_file, write_to_file, average_over=[67, 70],
                                      this_stand_only=67)


if __name__ == '__main__':
    test_load_fossagrim_input()