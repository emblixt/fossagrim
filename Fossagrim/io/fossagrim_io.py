import pandas as pd
import numpy as np
import os
import unittest

from Fossagrim.utils.definitions import heureka_mandatory_standdata_keys, \
    heureka_standdata_keys, heureka_standdata_desc, \
    heureka_treatment_keys, heureka_treatment_desc, \
    fossagrim_standdata_keys, \
    translate_keys_from_fossagrim_to_heureka


def get_row_index(table, key, item):
    """
    Returns the row index of the item
    :param table:
        panda DataFrame
    :param key:
        str
        Name of key we are searching within
    :param item:
        str
        The name we are looking for

    :return:
    """
    for i, var in enumerate(table[key]):
        if item in var:
            return i

    return None


def read_excel(filename, header, sheet_name):
    try:
        table = pd.read_excel(filename, header=header, sheet_name=sheet_name, engine='openpyxl')
    except PermissionError as err_msg:
        print(err_msg)
        return None
    return table


def write_csv_file(write_to_file, default_keys, default_desc, append=False, **kwargs):
    """
    Writes the parameter values from the kwargs to the 'write_to_file'.
    It writes to a semicolon separated data file whose first line contains a description, and second
    line the key names, then each line is a line of data

    Example output is the StandData.csv files that is used by Heureka for import of stand
    data

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
        if True, the data in kwargs are appended to an existing output file without creating a header
    :param kwargs:
        list of keyword arguments.
        keywords must match the list 'heureka_standdata_keys' given in utils.definitions.py
        The value of each key is used when writing the 'write_to_file' file
    :return:
    """
    if len(default_keys) != len(default_desc):
        raise IOError('The lists of keys and descriptions must be of same length')

    # overwrite any existing file, and create the header.
    if not append:
        with open(write_to_file, 'w') as f:
            f.write(';'.join(default_desc)+'\n')
            f.write(';'.join(default_keys)+'\n')

    # arrange the data from kwargs
    data = ['']*len(default_keys)
    for key in kwargs:
        if key not in default_keys:
            print('WARNING: key "{}" not found in accepted default keys'.format(key))
            continue
        this_data = kwargs[key]
        if isinstance(this_data, float):
            this_string = '{:.2f}'.format(this_data)
        else:
            this_string = str(this_data)
        data[default_keys.index(key)] = this_string

    # append the data
    if len(kwargs) > 0:  # only write data if given
        with open(write_to_file, 'a') as f:
            f.write(';'.join(data)+'\n')


def read_raw_heureka_results(filename, sheet_name):
    """
    Read heureka results as exported from Heureka (by simply copying a simulation result and pasting it
    into an empty sheet in Excel) and returns the data in a transposed DataFrame with one line for each period (5 years)
    NOTE, the Variable Treatments:Year must be included in the raw results

    :param filename:
    :param sheet_name:
    :return:
        panda DataFrame

    """
    table = read_excel(filename, 0, sheet_name)
    # Some treatments in Heureka break the default periods of 5 years into smaller sub-periods, we choose to remove
    # these from the returned dataframe
    year_row = get_row_index(table, 'Variable', 'Year')
    five_years = [np.mod(this_year, 5) == 0 for this_year in table.iloc[year_row, 3:]]
    # print(table.iloc[year_row, 3:][five_years])
    data_dict = {}
    for variable in list(table['Variable']):
        row_i = get_row_index(table, 'Variable', variable)
        data_dict[variable] = table.iloc[row_i, 3:][five_years]

    return pd.DataFrame(data=data_dict)


def rearrange_raw_heureka_results(filename, sheet_names):
    """

    :param filename:
    :param sheet_names:
        list
        List of strings with names of sheets that contains raw results from different Heureka simulations
    :return:
    """
    from openpyxl import load_workbook

    result_sheet_name = 'Rearranged results'

    this_start_col = 0
    start_cols = []
    for sheet_name in sheet_names:
        table = read_raw_heureka_results(filename, sheet_name)
        with pd.ExcelWriter(filename, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
            table.to_excel(writer, sheet_name=result_sheet_name, startcol=this_start_col, startrow=2)
        start_cols.append(this_start_col)
        this_start_col = len(list(table.keys())) + 2

    wb = load_workbook(filename)
    ws = wb[result_sheet_name]
    for i, header in enumerate(sheet_names):
        ws.cell(1, start_cols[i] + 1).value = header
    wb.save(filename)


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
        Dictionary of keyword: value pairs that can be used by 'write_csv_file'
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
        if key == 'N':
            keyword_arguments[key] = value * 10.  # trees/mål to trees/ha
        if key == 'PlantDensity':
            keyword_arguments[key] = value * 10.  # plants/daa to plants/ha
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
        if key in ['Fossagrim ID', 'Bestand']:
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
    See https://www.heurekaslu.se/wiki/Import_of_stand_register for description of parameters used by Heureka


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
    note = ''

    if (average_over is not None) and isinstance(average_over, list):  # export averaged stands
        this_stand_only = 'Avg stand'
        table = average_over_stands(average_over, table, stand_id_key, this_stand_only)
        note = 'Area weighted average of stand {}'.format(', '.join([str(_x) for _x in average_over]))

    for i, stand_id in enumerate(table[stand_id_key]):
        if (table['Bonitering\ntreslag'][i] == 'Uproduktiv') or np.isnan(table['HovedNr'][i]):
            continue
        if (this_stand_only is not None) and (this_stand_only != stand_id):
            continue
        # get the data from the fossagrim table arranged for writing in heureka format
        keyword_arguments = arrange_export_of_one_stand(i, table)
        write_csv_file(
            write_to_file,
            heureka_standdata_keys,
            heureka_standdata_desc,
            append=append,
            Note=note,
            **keyword_arguments
        )
        append = True


def export_fossagrim_treatment(read_from_file, write_to_file, this_stand_only=None,
                               average_over=None, stand_id_key=None, header=None, sheet_name=None):
    """
    Load a 'typical' Fossagrim stand data file (Bestandsutvalg) and writes a 'treatment proposal' output file.
    This treatment proposal file does not fully follow the Heureka standard for TreatmentProposals but contains Fossagrim
    specific changes
    See https://www.heurekaslu.se/wiki/Import_of_stand_register for description of parameters used by Heureka
    and https://www.heurekaslu.se/help/index.htm?importera_atgardsforslag.htm for treatment proposals


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

    # write to file, but only header lines
    write_csv_file(
        write_to_file,
        heureka_treatment_keys,
        heureka_treatment_desc,
        append=False
    )

    note = ''
    if (average_over is not None) and isinstance(average_over, list):  # export averaged stands
        this_stand_only = 'Avg stand'
        table = average_over_stands(average_over, table, stand_id_key, this_stand_only)
        note = 'Area weighted average of stand {}'.format(', '.join([str(_x) for _x in average_over]))

    for i, stand_id in enumerate(table[stand_id_key]):
        if (table['Bonitering\ntreslag'][i] == 'Uproduktiv') or np.isnan(table['HovedNr'][i]):
            continue
        if (this_stand_only is not None) and (this_stand_only != stand_id):
            continue
        # Write the final felling at year 0 with subsequent planting
        write_csv_file(
            write_to_file, heureka_treatment_keys, heureka_treatment_desc, append=True,
            StandId=table['Fossagrim ID'][i], Year=0, Treatment='FinalFeeling'
        )
        write_csv_file(
            write_to_file, heureka_treatment_keys, heureka_treatment_desc, append=True,
            StandId=table['Fossagrim ID'][i], Year=2, Treatment='Planting', PlantDensity=table['Plantetetthet'][i] * 10.
        )
        # write the thinning
        write_csv_file(
            write_to_file, heureka_treatment_keys, heureka_treatment_desc, append=True,
            StandId=table['Fossagrim ID'][i], Year=table['Tynnings år'][i], Treatment='Thinning'
        )
        # And the next final felling with subsequent planting
        write_csv_file(
            write_to_file, heureka_treatment_keys, heureka_treatment_desc, append=True,
            StandId=table['Fossagrim ID'][i], Year=table['Rotasjonsperiode'][i], Treatment='FinalFeeling'
        )
        write_csv_file(
            write_to_file, heureka_treatment_keys, heureka_treatment_desc, append=True,
            StandId=table['Fossagrim ID'][i], Year=table['Rotasjonsperiode'][i] + 2,
            Treatment='Planting', PlantDensity=table['Plantetetthet'][i] * 10.
        )


def test_rearrange():
    file = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-003 Kloppmyra\\FHF23-003 Heureka results COPY.xlsx"
    rearrange_raw_heureka_results(file, ['FHF23-003 Business as usual', 'FHF23-003 Preservation'])


def test_write_csv_file():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    write_csv_file(os.path.join(dir_path, 'test.csv'),
                   heureka_standdata_keys,
                   heureka_standdata_desc,
                   test=8, DGV=46, SiteIndexSpecies='G')
    write_csv_file(os.path.join(dir_path, 'test.csv'),
                   heureka_standdata_keys,
                   heureka_standdata_desc,
                   append=True, SiteIndexSpecies='T')
    write_csv_file(os.path.join(dir_path, 'test2.csv'),
                   heureka_standdata_keys,
                   heureka_standdata_desc)

    unittest.TestCase.assertTrue(True)


def test_export_fossagrim_stand():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    read_from_file = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-003 Kloppmyra\\FHF23-003 Bestandsutvalg.xlsx"
    write_to_file = os.path.join(dir_path, 'standdata.csv')

    export_fossagrim_stand_to_heureka(read_from_file, write_to_file, average_over=[67, 70],
                                      this_stand_only=67)


def test_export_fossagrim_treatment():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    read_from_file = "C:\\Users\\marten\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF23-003 Kloppmyra\\FHF23-003 Bestandsutvalg.xlsx"
    write_to_file = os.path.join(dir_path, 'treatment.csv')

    export_fossagrim_treatment(read_from_file, write_to_file, average_over=[67, 70],
                               this_stand_only=67)


if __name__ == '__main__':
    test_rearrange()