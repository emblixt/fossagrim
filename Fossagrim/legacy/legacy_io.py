"""
LEGACY functions that used to live in Fossagrim/io/fossagrim_io.py.

They are no longer used by the active workflow (create_heureka_input_files.py, rearrange_heureka_results.py,
qc_input_pdfs.py, utils/projects_new.py), and are kept for reference only. They are not covered by unit tests.
"""
import os

import matplotlib.pyplot as plt
import numpy as np
import openpyxl

import Fossagrim.plotting.misc_plots as fpp
from Fossagrim.io.common import read_excel, write_csv_file
from Fossagrim.io.stands import average_over_stands
from Fossagrim.utils.definitions import heureka_treatment_keys, heureka_treatment_desc


def example_gis_database():
    """
    This is just a note on how to start handling tables from gis (exported from Allma)
    :return:
    """
    pass
    # import geopandas as gpd
    # # TODO

    # base_dir = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\Sandbox\\Allma\\TEST Nord-Odal\\04180001900020000.gdb"

    # # The following loop goes through all data tables in the above gis project and locates the table that contains
    # # the key 'BEST_NR'
    # for filename in os.listdir(base_dir):
    #     if "gdbtable" in filename:
    #         print('Reading {}'.format(filename))
    #         data = gpd.read_file(os.path.join(base_dir, filename))
    #         if 'BEST_NR' in list(data.keys()):
    #             break
    # # And the following line finds the index of bestand nr. 279
    # index = np.where(data['BEST_NR'] == 279.0)



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



def export_fossagrim_treatment(read_from_file, write_to_file, this_stand_only=None,
                               average_over=None, stand_id_key=None, header=None, sheet_name=None, average_name=None):
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
        dict
        dictionary where each item contains a list of strings/int which determines which stands to use in calculating
        average stands. The key is the name of each individual average. eg.
        {
            'Spruce': [67, 68, 70, 72, 75],
            'Pine': [32, 56, 58]
        }
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
        sheet_name = 0  # Reads only first sheet, opposite to pandas default were None reads all sheets.
    if header is None:
        header = 7
    if average_name is None:
        average_name = 'Avg Stand'

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

    notes = [''] * len(table[stand_id_key])

    if (average_over is not None) and isinstance(average_over, dict):  # export averaged stands
        this_stand_only = None
        table = average_over_stands(average_over, table, stand_id_key, average_name)
        notes = [
            'Area weighted average for {} of stands {}'.format(_key, ', '.join([str(_x) for _x in average_over[_key]]))
            for _key in list(average_over.keys())]

    for i, stand_id in enumerate(table[stand_id_key]):
        if (table['Bonitering\ntreslag'][i] in ['Uproduktiv', '-']):  # or np.isnan(table['HovedNr'][i]):
            continue
        if (this_stand_only is not None) and (this_stand_only != stand_id):
            continue
        # Write the final felling at year 0 with subsequent planting
        write_csv_file(
            write_to_file, heureka_treatment_keys, heureka_treatment_desc, append=True,
            StandId=table['Fossagrim ID'][i], Year=0,
            Treatment='FinalFelling', Note=notes[i]
        )
        write_csv_file(
            write_to_file, heureka_treatment_keys, heureka_treatment_desc, append=True,
            StandId=table['Fossagrim ID'][i], Year=2,
            Treatment='Planting', PlantDensity=table['Plantetetthet'][i] * 10., Note=notes[i]
        )
        # write the thinning
        write_csv_file(
            write_to_file, heureka_treatment_keys, heureka_treatment_desc, append=True,
            StandId=table['Fossagrim ID'][i], Year=table['Tynnings år'][i],
            Treatment='Thinning', Note=notes[i]
        )
        # And the next final felling with subsequent planting
        write_csv_file(
            write_to_file, heureka_treatment_keys, heureka_treatment_desc, append=True,
            StandId=table['Fossagrim ID'][i], Year=table['Rotasjonsperiode'][i],
            Treatment='FinalFelling', Note=notes[i]
        )
        write_csv_file(
            write_to_file, heureka_treatment_keys, heureka_treatment_desc, append=True,
            StandId=table['Fossagrim ID'][i], Year=table['Rotasjonsperiode'][i] + 2,
            Treatment='Planting', PlantDensity=table['Plantetetthet'][i] * 10., Note=notes[i]
        )



def read_fossagrim_treatment(
        treatment_csv_file,
        project_tag,
        wood_species='Spruce'):
    """

    :param treatment_csv_file:
    :param project_tag:
    :param wood_species:
    :return:
        list
        [thinning_year, final_felling_year, plant_density]  # [year, year, plants/ha]
    """
    thinning_year = None
    final_felling_year = None
    plant_density = None
    with open(treatment_csv_file) as f:
        lines = f.readlines()
        for line in lines:
            split_line = line.split(';')
            if '{} {}'.format(project_tag, wood_species) == split_line[2]:
                if 'FinalFelling' in line:
                    final_felling_year = float(split_line[4])
                if 'Thinning' in line:
                    thinning_year = float(split_line[4])
                if 'Planting' in line:
                    plant_density = float(split_line[8])
    return [thinning_year, final_felling_year, plant_density]



def get_kwargs_from_stand_OLD(stand_file, project_settings_file, project_tag):
    """
    Extracts necessary information from stand_file and project_settings_file to feed modify_monetization_file()
    :param stand_file:
        str
        Full path name of Fossagrim stand data file (Bestandsutvalg)
    :param project_settings_file:
        str
        Full path name of the project settings Excel sheet.
    :param project_tag:
        str
        Name tag that identifies the current project, e.g. "FHF23-007"
    :return:
        dict, list
        kwargs: Dictionary of kwargs needed by modify_monetization_file()
        coombine_fractions: list of combination fractions read from the stand_file
    """
    p_tabl = read_excel(project_settings_file, 1, 'Settings')
    i = None
    # Extract the project settings for the given project
    for i, p_name in enumerate(p_tabl['Project name']):
        if not isinstance(p_name, str):
            continue
        if project_tag in p_name:
            break

    combine_positions = [_x.strip() for _x in p_tabl['Position of fractions when combined'][i].split(',')]
    wb = openpyxl.load_workbook(stand_file, data_only=True)
    ws = wb.active
    combine_fractions = [ws[_x].value for _x in combine_positions]
    _kwarg_position_keys = [
        'Position of total area',
        'Position of Batch 1 total volume',
        'Position of Batch 2 total volume',
        'Position of passive forest total volume',
    ]
    _kwarg_direct_keys = [
        'Batch 1 start date',
        'Batch 2 delay',
        'Root net',
        'Contract length',
        'Rent',
        'Price growth',
        'Buffer',
        'Reserve years',
        'Net price',
        'Gross price'
    ]
    kwargs = {
        _key: ws[p_tabl[_key][i]].value for _key in _kwarg_position_keys
    }
    for _key in _kwarg_direct_keys:
        kwargs[_key] = p_tabl[_key][i]
    wb.close()

    return kwargs, combine_fractions


def get_carbon_effect(result_file, stand_id,
                      average_years=None,
                      variable="Total Carbon Stock (dead wood, soil, trees, stumps and roots)",
                      ax=None,
                      save_plot_to=None):
    """
    Calculates the difference in Total carbon stock between PRES (preservation) and BAU (Business as usual)
    integrated over the given number of years
    :param result_file:
         str
         Name of excel file that contain the "raw" Heureka results
    :param stand_id
         str
    :param average_years:
         int
         Number of years to calculate the integrated effect over
    :param variable:
        str
        Variable name to use in the calculation of the carbon effect
    :param save_plot_to:
        str
        Name of file to save the plot to.
    :return:
        float
        Integrated effect of chosen parameter (variable) over <average_years> years for the given forest stand
        (project_tag) and wood_species
    """
    if average_years is None:
        average_years = 30
    try:
        diff = fpp.plot_from_heureka_results(
            result_file=result_file,
            sheets=['{} {}'.format(stand_id, _case) for _case in ['PRES', 'BAU']],
            params=variable,
            ax=ax,
            save_plot_to=save_plot_to,
            diff_sheets=True
        )
        # Heureka runs the simulation using timesteps which are 5 years long
        i = int(np.floor(average_years / 5))
        if diff is not None:
            return np.sum(diff[:i])
        else:
            return None
    except ValueError as error:
        print(error)
        return None



def collect_all_stand_data_OLD(
        out_file=None,
        base_dir=None,
        dry_run=True
):
    """
    Traverses the file structure for Heureka files and collects the stand data and selected Heureka
     results in one csv file

    out_file:
        str
        Full file name of output file
    :param base_dir:
        Path to directory where Heureka modelling results are stored (in sub folders of base_dir)
    :param dry_run:
        bool
        If True, not results are written, just checking for potential Heureka files
    :return:
    """
    if base_dir is None:
        base_dir = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger"

    if dry_run:
        # Create axes for plotting nothing
        fig, ax = plt.subplots()
    else:
        # get_carbon_effect() creates a new plot for each stand
        ax = None

    from pathlib import Path
    if dry_run and out_file is None:
        out_file = os.path.join(base_dir, 'TEST COLLECT STAND DATA.csv')

    if out_file is None:
        raise IOError('Name of output file must be given')

    file_list = Path(base_dir).rglob('*Averaged stand data.csv')
    with open(out_file, 'a') as _out:  # append
        for stand_file in file_list:
            print('Working on {}'.format(stand_file.name))
            # see if there is a related Heureka results file
            heureka_result_file = str(stand_file).replace('Averaged stand data.csv', 'Heureka results.xlsx')
            if os.path.isfile(heureka_result_file):
                print(' Heureka result file exists')
            else:
                continue

            with open(stand_file, 'r') as _in:
                for line in _in.readlines():
                    split_line = line.split(';')
                    if 'FHF' in split_line[0]:
                        print('  Searching for stand id: {}'.format(split_line[0]))
                        if dry_run:
                            save_plot = None
                        else:
                            save_plot = os.path.join(
                                os.path.dirname(heureka_result_file),
                                'Net Carbon Effect {}.png'.format(split_line[0]))

                        carbon_effect = get_carbon_effect(
                            heureka_result_file,
                            split_line[0],
                            ax=ax,
                            save_plot_to=save_plot
                        )
                        if carbon_effect is None:
                            alt_stand_id = split_line[0].replace('Avg Stand-', '')
                            print('  Searching for alternative stand id: {}'.format(alt_stand_id))
                            carbon_effect = get_carbon_effect(
                                heureka_result_file,
                                alt_stand_id,
                                ax=ax,
                                save_plot_to=save_plot
                            )

                        print('  ', carbon_effect)
                        split_line[101] = str(carbon_effect)
                        if not dry_run:
                            _out.write(';'.join(split_line))
                        else:
                            print('   - Dry run')

