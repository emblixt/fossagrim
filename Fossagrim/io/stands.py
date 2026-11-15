"""
Reading stands from the Fossagrim "Forvaltningsplan" (a.k.a. "Bestandsutvalg") Excel file, averaging them
per tree species, and exporting them to a stand register csv file that Heureka can import.

See https://www.heurekaslu.se/wiki/Import_of_stand_register for the parameters used by Heureka.

Typical flow (see create_heureka_input_files.py):
    get_active_forests_from_stand()      -> which stands are active, grouped per species
    export_fossagrim_stand_to_heureka()  -> average_over_stands() + arrange_export_of_one_stand() + write_csv_file()
"""
import numpy as np
import pandas as pd

from Fossagrim.io.common import my_float, read_excel, write_csv_file
from Fossagrim.utils.definitions import (
    heureka_mandatory_standdata_keys, heureka_standdata_keys, heureka_standdata_desc,
    fossagrim_keys_to_average_over, fossagrim_keys_to_sum_over, fossagrim_keys_to_most_of,
    m3fub_to_m3sk, translate_keys_from_fossagrim_to_heureka,
    FORVALTNINGSPLAN_DATA_SHEET, FORVALTNINGSPLAN_HEADER_ROW, SPECIES_KEY, UNPRODUCTIVE_SPECIES, SPECIES_NAMES,
    HEUREKA_SITE_INDEX_SPECIES)

# Non-default Heureka keys that are always exported in addition to heureka_mandatory_standdata_keys
EXTRA_EXPORT_KEYS = ['PlantDensity', 'UserDefinedVariable2_RotationPeriod', 'UserDefinedVariable3_ThinningYear']

DEFAULT_AVERAGE_NAME = 'Avg Stand'


def arrange_export_of_one_stand(row_index, table, extra_keywords_list=None, unique_id=None):
    """
    Collects, translates, and converts the Fossagrim data of one stand into the keys and units used by Heureka

    Unit conversions:
        ProdArea:       daa -> ha
        V:              m3fub/daa -> m3sk/ha
        N:              trees/daa -> trees/ha
        PlantDensity:   plants/daa -> plants/ha
        PropPine/PropSpruce/PropBirch: volume fraction of Furu/Gran/Lauv relative to Total
        CountyCode:     always 13

    :param row_index:
        int
        row index (beginning at 0) of table which we want to export
    :param table:
        pandas DataFrame
        Forvaltningsplan table (as read with read_excel, or the output of average_over_stands)
        Each row contains one stand
    :param extra_keywords_list:
        list
        List of non-default keywords from heureka_standdata_keys that we want to add
    :param unique_id:
        str
        When not None, this string is stored in the column "UniqueID" (column DA when opened in Excel).
        It is used as an identifier in the Woods platform to allow us to separate between different versions.
        Only used if 'UniqueID' is in extra_keywords_list.
    :return:
        dict
        Dictionary of keyword: value pairs that can be used by 'write_csv_file'
    """
    key_translation = translate_keys_from_fossagrim_to_heureka()
    keyword_arguments = {}

    if extra_keywords_list is None:
        extra_keywords_list = []

    # Collect the raw values (no unit conversion yet) for all keys that Heureka needs
    for key in heureka_mandatory_standdata_keys + extra_keywords_list:
        keyword_arguments[key] = ''
        if key in key_translation:
            keyword_arguments[key] = table[key_translation[key]][row_index]

    total_volume = table['Total'][row_index]

    # Convert from Fossagrim data and units to Heureka data and units
    for key, value in list(keyword_arguments.items()):
        if key == 'CountyCode':
            keyword_arguments[key] = 13
        elif key in ['InventoryYear', 'SoilMoistureCode', 'VegetationType', 'Peat']:
            try:
                keyword_arguments[key] = int(value)
            except ValueError as err_msg:
                print('For key {}: {}'.format(key, err_msg))
        elif key == 'ProdArea':
            keyword_arguments[key] = value * 0.1  # from daa to ha
        elif key == 'SiteIndexSpecies':
            if value not in HEUREKA_SITE_INDEX_SPECIES:
                raise NotImplementedError('Species "{}" is not recognized as one of {}'.format(
                    value, ', '.join(HEUREKA_SITE_INDEX_SPECIES.keys())))
            keyword_arguments[key] = HEUREKA_SITE_INDEX_SPECIES[value]
        elif key == 'V':
            keyword_arguments[key] = value * 10. * m3fub_to_m3sk  # m3fub/daa to m3sk/ha
        elif key == 'N':
            keyword_arguments[key] = value * 10.  # trees/daa to trees/ha
        elif key == 'PlantDensity':
            keyword_arguments[key] = value * 10.  # plants/daa to plants/ha
        elif key == 'PropPine':
            keyword_arguments[key] = table['Furu'][row_index] / total_volume
        elif key == 'PropSpruce':
            keyword_arguments[key] = table['Gran'][row_index] / total_volume
        elif key == 'PropBirch':
            keyword_arguments[key] = table['Lauv'][row_index] / total_volume
        elif key == 'StandId':
            keyword_arguments[key] = table['Fossagrim ID'][row_index]
        elif key == 'UniqueID':
            keyword_arguments[key] = unique_id

    return keyword_arguments


def _rows_in_group(table, stand_id_key, stand_ids):
    """Row indexes of 'table' whose 'stand_id_key' value is in the list 'stand_ids'"""
    return [i for i, stand_id in enumerate(table[stand_id_key]) if stand_id in stand_ids]


def matched_average_groups(average_over, table, stand_id_key):
    """
    The keys of 'average_over' that match at least one stand in 'table', in the same order as the rows
    that average_over_stands() creates.
    """
    return [group for group in average_over if len(_rows_in_group(table, stand_id_key, average_over[group])) > 0]


def average_over_stands(average_over, table, stand_id_key, average_name, verbose=False):
    """
    Calculates one area weighted "average stand" for each group of stands.

    How each column is treated is decided by the lists in utils/definitions.py:
        'Fossagrim ID', 'Bestand':           set to '<average_name><group name>'
        fossagrim_keys_to_sum_over:          sum over the stands
        fossagrim_keys_to_average_over:      area ('Prod.areal') weighted average
        fossagrim_keys_to_most_of:           the value of the stand that is closest to the area weighted average
        all other columns:                   copied from the first stand in the group

    :param average_over:
        dict
        dictionary where each item contains a list of strings/int which determines which stands to use in calculating
        average stands. The key is the name of each individual average. eg.
        {
            'Spruce': [67, 68, 70, 72, 75],
            'Pine': [32, 56, 58]
        }
    :param table:
        pandas DataFrame
        Forvaltningsplan table as read with read_excel. Each row contains one stand
    :param stand_id_key:
        str
        Key that is used to identify the stands used in 'average_over'
    :param average_name:
        str
        Given main name of the calculated average stands. The key in 'average_over' is then added, e.g.
        '<average_name>Spruce'
    :param verbose:
        bool
    :return:
        pandas DataFrame
        Same columns as the input table, with one row for each group in 'average_over' that matched at least
        one stand (groups without any matching stands are skipped with a warning)
    """
    avg_table = pd.DataFrame(columns=[_key.strip() for _key in table.keys()])

    for avg_group in average_over:
        if verbose:
            print('Averaging group {}, which contains stands {}'.format(avg_group, average_over[avg_group]))
        average_ind = _rows_in_group(table, stand_id_key, average_over[avg_group])

        if len(average_ind) == 0:
            print('WARNING, no match for average group {}'.format(avg_group))
            continue

        area = np.array([my_float(_x) for _x in table['Prod.areal'].iloc[average_ind]])
        total_area = np.sum(area)

        this_row_of_data = []
        for key in table.keys():
            column = table[key].iloc[average_ind]
            if key.strip() in ['Fossagrim ID', 'Bestand']:
                this_data = '{}{}'.format(average_name, avg_group)
                rule = 'Name'
            elif key.strip() in fossagrim_keys_to_sum_over:
                this_data = np.sum([my_float(_x) for _x in column])
                rule = 'Sum over'
            elif key.strip() in fossagrim_keys_to_average_over:
                this_data = np.sum(area * np.array([my_float(_x) for _x in column])) / total_area
                rule = 'Average over'
            elif key.strip() in fossagrim_keys_to_most_of:
                # An area weighted 'most of': pick the original value closest to the area weighted average
                values = np.array([my_float(_x) for _x in column])
                area_weighted_average = np.sum(values * area) / total_area
                ind = np.argmin((values - area_weighted_average) ** 2)
                this_data = column.iloc[ind]
                rule = 'Most of'
            else:
                this_data = column.iloc[0]
                rule = 'Copy first'
            if verbose:
                print(' {}: {}: {}'.format(rule, key, this_data))
            this_row_of_data.append(this_data)

        avg_table.loc[len(avg_table)] = this_row_of_data

    return avg_table


def export_fossagrim_stand_to_heureka(read_from_file, write_to_file, this_stand_only=None, average_over=None,
                                      stand_id_key=None, header=None, sheet_name=None, average_name=None,
                                      verbose=False, unique_id=None):
    """
    Load a Fossagrim stand data file (Forvaltningsplan / Bestandsutvalg) and write a stand register csv file
    which Heureka can import.
    See https://www.heurekaslu.se/wiki/Import_of_stand_register for description of parameters used by Heureka

    Unproductive stands (see UNPRODUCTIVE_SPECIES in utils/definitions.py) are never exported.

    :param read_from_file:
        str
        Full path name of the Forvaltningsplan Excel file
    :param write_to_file:
        str
        Full path name of the csv file to write. Overwritten if it exists.
    :param this_stand_only:
        str
        Name of the only stand to export. Ignored when 'average_over' is given
    :param average_over:
        dict
        When given, the area weighted average stands are exported instead of the individual stands.
        See average_over_stands() for the format, eg.
        {
            'Spruce': [67, 68, 70, 72, 75],
            'Pine': [32, 56, 58]
        }
    :param stand_id_key:
        str
        Key that is used to identify the stands used in 'average_over' and 'this_stand_only'. Default 'Bestand'
    :param header:
        int
        Zero-based row index of the column names. Default FORVALTNINGSPLAN_HEADER_ROW
    :param sheet_name:
        str or int
        Default: the first sheet
    :param average_name:
        str
        Prefix of the names of the average stands, see average_over_stands(). Default 'Avg Stand'
    :param verbose:
        bool
    :param unique_id:
        str or int
        When not None, it is stored in column "UniqueID" (column DA when opened in Excel) for each exported stand.
        It is used as an identifier in the Woods platform to allow us to separate between different versions
    :return:
        int
        Number of stands written, or None if the Forvaltningsplan could not be read
    """
    if stand_id_key is None:
        stand_id_key = 'Bestand'
    if sheet_name is None:
        sheet_name = 0  # Reads only first sheet, opposite to pandas default were None reads all sheets.
    if header is None:
        header = FORVALTNINGSPLAN_HEADER_ROW
    if average_name is None:
        average_name = DEFAULT_AVERAGE_NAME

    table = read_excel(read_from_file, header, sheet_name)
    if table is None:
        print('WARNING, stands could not be loaded from {}'.format(read_from_file))
        return None

    notes = [''] * len(table[stand_id_key])

    if isinstance(average_over, dict):  # export averaged stands
        this_stand_only = None
        # The notes must be built from the groups that actually ended up in the averaged table
        groups = matched_average_groups(average_over, table, stand_id_key)
        table = average_over_stands(average_over, table, stand_id_key, average_name, verbose=verbose)
        notes = [
            'Area weighted average for {} of stands {}'.format(_key, ', '.join([str(_x) for _x in average_over[_key]]))
            for _key in groups]

    extra_keywords_list = list(EXTRA_EXPORT_KEYS)
    if unique_id is not None:
        extra_keywords_list.append('UniqueID')

    append = False
    number_of_stands_written = 0
    for i, stand_id in enumerate(table[stand_id_key]):
        if table[SPECIES_KEY][i] in UNPRODUCTIVE_SPECIES:
            if verbose:
                print('Skipping unproductive stand {}'.format(stand_id))
            continue
        if (this_stand_only is not None) and (this_stand_only != stand_id):
            continue
        if verbose:
            print('Attempting to write {} to {}'.format(stand_id, write_to_file))
        keyword_arguments = arrange_export_of_one_stand(i, table, extra_keywords_list, unique_id=unique_id)
        write_csv_file(
            write_to_file,
            heureka_standdata_keys,
            heureka_standdata_desc,
            append=append,
            Note=notes[i],
            **keyword_arguments
        )
        number_of_stands_written += 1
        append = True

    if number_of_stands_written == 0:
        print('WARNING: No stands written to {}'.format(write_to_file))

    return number_of_stands_written


def get_active_forests_from_stand(stand_file, stand_id_key=None, header=None, sheet_name=None):
    """
    Reads the Forvaltningsplan and finds the active stands (column 'Volum status' == 1), groups them by their
    dominant tree species, and calculates the productive area of each species group

    :param stand_file:
        str
        Full path name of the Forvaltningsplan Excel file
    :param stand_id_key:
        str
        Column used to identify the stands. Default 'Bestand'
    :param header:
        int
        Default FORVALTNINGSPLAN_HEADER_ROW
    :param sheet_name:
        str or int
        Default: the first sheet
    :return:
        dict, dict, dict   (or None if the file could not be read)
        active_forests:              {'Spruce': [stand ids], 'Pine': [...], ...}
        active_prod_areal:           {'Spruce': [productive area (daa) of each stand], ...}
        active_prod_areal_fractions: {'Spruce': fraction of total active area, ...}
        Species without any active stands are left out.
    """
    if stand_id_key is None:
        stand_id_key = 'Bestand'
    if sheet_name is None:
        sheet_name = 0  # Reads only first sheet, opposite to pandas default were None reads all sheets.
    if header is None:
        header = FORVALTNINGSPLAN_HEADER_ROW

    table = read_excel(stand_file, header, sheet_name)
    if table is None:
        print('WARNING, stands could not be loaded from {}'.format(stand_file))
        return None

    active_forests = {species: [] for species in SPECIES_NAMES}
    active_prod_areal = {species: [] for species in SPECIES_NAMES}  # Unit daa in Forvaltningsplan

    for i, stand_id in enumerate(table[stand_id_key]):
        if table['Volum status'][i] != 1:
            continue
        this_wood_species = None
        for species, names in SPECIES_NAMES.items():
            if table[SPECIES_KEY][i] in names:
                this_wood_species = species
        if this_wood_species is None:
            raise ValueError('Active stand {} has an unrecognized tree species "{}" in column "{}"'.format(
                stand_id, table[SPECIES_KEY][i], SPECIES_KEY.replace('\n', ' ')))
        active_forests[this_wood_species].append(stand_id)
        active_prod_areal[this_wood_species].append(my_float(table['Prod.areal'][i]))

    # Remove wood species without active stands
    for key in list(active_forests.keys()):
        if len(active_forests[key]) == 0:
            active_forests.pop(key)
            active_prod_areal.pop(key)

    tot_area_per_species = {key: sum(value) for key, value in active_prod_areal.items()}
    tot_area = sum(tot_area_per_species.values())
    if tot_area <= 0:
        raise ValueError('No active stands with a positive productive area found in {}'.format(stand_file))
    active_prod_areal_fractions = {key: value / tot_area for key, value in tot_area_per_species.items()}

    return active_forests, active_prod_areal, active_prod_areal_fractions
