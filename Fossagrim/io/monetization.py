"""
The monetization Excel file: climate benefit and money flow calculations based on the rearranged Heureka results.

    get_kwargs_from_stand()     reads the project constants (areas, volumes, prices, ...) needed by
                                modify_monetization_file() from the Forvaltningsplan and the project settings file
    modify_monetization_file()  writes the formulas (see utils/monetization_parameters.py) and constants
    style_monetization_file()   colors, borders and number formats
    qc_plots()                  plots from a monetization file that has been opened and saved in Excel
                                (so that the formulas have calculated values)

NOTE: The monetization is currently done directly in Excel, and this module is only used by the legacy
workflow in legacy/projects.py.
"""
import os

import matplotlib.pyplot as plt
import numpy as np
import openpyxl
import pandas as pd
from openpyxl.workbook.defined_name import DefinedName

from Fossagrim.io.common import read_excel
from Fossagrim.utils.monetization_parameters import (
    parameters, calculation_part1, resampled_section, money_value, cbo_flow, project_benefits, buffer,
    fossagrim_values, forest_owner_values)


def get_kwargs_from_stand(stand_file, project_settings_file, project_tag):
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
        combine_fractions: list of combination fractions read from the stand_file
    """
    p_tabl = read_excel(project_settings_file, 1, 'Settings')
    # Extract the project settings for the given project
    i = None
    for row, p_name in enumerate(p_tabl['Project name']):
        if isinstance(p_name, str) and project_tag in p_name:
            i = row
            break

    if i is None:
        raise IOError('Project tag {} not found in {}'.format(
            project_tag, os.path.basename(project_settings_file)
        ))

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
    _kwarg_direct_keys = {
        'Batch 1 start date': 'J',
        'Batch 2 delay': 'L',
        'Root net': 'N',
        'Contract length': 'O',
        'Rent': 'P',
        'Price growth': 'Q',
        'Buffer': 'R',
        'Reserve years': 'S',
        'Net price': 'T',
        'Gross price': 'U',
        'NIBOR 10yr': 'V'
    }

    try:
        kwargs = {
            _key: ws[p_tabl[_key][i]].value for _key in _kwarg_position_keys
        }
    except TypeError as error:
        print(error)
        print(i, _kwarg_position_keys)
        for _key in _kwarg_position_keys:
            print(_key, p_tabl[_key][i])
        raise TypeError(error)

    for _key in list(_kwarg_direct_keys.keys()):
        # kwargs[_key] = "='C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\[ProjectForestsSettings – WIP.xlsx]Settings'!${}${}".format(
        #     _kwarg_direct_keys[_key], i+3)  # +3 because 2 header lines and python starts counting at 0
        # Using the absolut path makes the linking works easier, but makes it difficult to share the file
        # UPDATE! If I open the linked file (ProjectForestSettings - WIP.xlsx) the linking works:-)
        kwargs[_key] = "='[{}]Settings'!${}${}".format(
            os.path.basename(project_settings_file),
            _kwarg_direct_keys[_key], i + 3)  # +3 because 2 header lines and python starts counting at 0
    wb.close()

    if None in combine_fractions:
        raise IOError('Some fractions are lacking in cells: {} in {}'.format(
            ', '.join(combine_positions), stand_file
        ))

    return kwargs, combine_fractions


def modify_monetization_file(write_to_file, **_kwargs):
    """
    Writes the formulas of the monetization calculation, and the project constants in _kwargs, into an
    existing monetization file that contains the sheets 'Parameters' and 'Monetization' (see
    heureka_results.rearrange_raw_heureka_results()), styles it, and creates the named cells used in the formulas.

    :param write_to_file:
        str
    :param _kwargs:
        The project constants as returned by get_kwargs_from_stand()
    """

    with pd.ExcelWriter(write_to_file, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
        parameters.to_excel(
            writer, sheet_name='Parameters', startcol=0, startrow=1, index=False, header=False)
        calculation_part1.to_excel(
            writer, sheet_name='Monetization', startcol=0, startrow=0, index=False, header=False)
        resampled_section.to_excel(
            writer, sheet_name='Monetization', startcol=24, startrow=0, index=False, header=False)
        money_value.to_excel(
            writer, sheet_name='Monetization', startcol=30, startrow=1, index=False, header=False)
        cbo_flow.to_excel(
            writer, sheet_name='Monetization', startcol=35, startrow=1, index=False, header=False)
        project_benefits.to_excel(
            writer, sheet_name='Monetization', startcol=39, startrow=1, index=False, header=False)
        buffer.to_excel(
            writer, sheet_name='Monetization', startcol=49, startrow=2, index=False, header=False)
        fossagrim_values.to_excel(
            writer, sheet_name='Monetization', startcol=55, startrow=1, index=False, header=False)
        forest_owner_values.to_excel(
            writer, sheet_name='Monetization', startcol=66, startrow=1, index=False, header=False)

    # Modify monetization file with constants kwargs
    wb = openpyxl.load_workbook(write_to_file)

    ws = wb['Monetization']
    ws['F1'].value = _kwargs['Position of total area'] / 10.  # from mål to Ha
    ws['AF3'].value = _kwargs['Position of Batch 1 total volume']
    ws['AG3'].value = _kwargs['Position of Batch 2 total volume']
    ws['AI3'].value = _kwargs['Position of Batch 1 total volume'] + _kwargs['Position of Batch 2 total volume'] + \
                      _kwargs['Position of passive forest total volume']
    ws['AF5'].value = _kwargs['Batch 1 start date']
    ws['AG6'].value = _kwargs['Batch 2 delay']
    ws['AF8'].value = _kwargs['Root net']
    ws['AN4'].value = _kwargs['Contract length']
    ws['AO4'].value = _kwargs['Rent']
    ws['AP4'].value = _kwargs['Price growth']
    ws['AX4'].value = _kwargs['Buffer']
    ws['AY4'].value = _kwargs['Reserve years']
    ws['BH8'].value = _kwargs['Net price']
    ws['BI8'].value = _kwargs['Gross price']
    ws['BP3'].value = _kwargs['NIBOR 10yr']

    wb.save(write_to_file)

    style_monetization_file(write_to_file)

    # Create defined name (name manager) using openpyxl
    wb = openpyxl.load_workbook(write_to_file)
    defined_names = {
        # name: cell in the Parameters sheet, see 'parameters' in utils/monetization_parameters.py
        'hpap': 'B2', 'hsawn': 'B3', 'SFsawn': 'B4', 'SFpp': 'B5', 'SFfuel': 'B6',
        'p': 'B7', 'pp': 'B8', 'pf': 'B9', 'psfuel': 'B10', 'k': 'B11'}
    for name, cell in defined_names.items():
        col, row = cell[0], cell[1:]
        wb.defined_names.add(DefinedName(name, attr_text='Parameters!${}${}'.format(col, row)))
    wb.save(write_to_file)


def style_monetization_file(write_to_file):
    """Colors, borders, merged cells and number formats of the Monetization sheet"""
    from openpyxl.styles import NamedStyle
    from openpyxl.styles import PatternFill
    from openpyxl.styles import Alignment
    from openpyxl.styles import Font
    from Fossagrim.utils.definitions import standard_colors as scrs

    colors = {
        scrs['time_1'].replace('#', ''): ['A1:G50'],
        scrs['products_1'].replace('#', ''): ['C3:F50'],
        scrs['substitution_1'].replace('#', ''): ['G3:G50'],
        scrs['bau_1'].replace('#', ''): ['H1:O50'],
        scrs['bau_2'].replace('#', ''): ['L2:O50'],
        scrs['project_case_1'].replace('#', ''): ['P1:Q50', 'BO2:BT110'],
        scrs['project_case_2'].replace('#', ''): ['Q2:Q50', 'BD2:BN110'],
        scrs['climate_benefit_1'].replace('#', ''): ['R1:W50', 'AN2:AW110'],
        scrs['climate_benefit_2'].replace('#', ''): ['U2:W50', 'AC1:AD110'],
        scrs['buffer_1'].replace('#', ''): ['Y1:AB110', 'AX2:BC110'],
        scrs['table_1'].replace('#', ''): ['AE2:AI9'],
        scrs['cbo_1'].replace('#', ''): ['AJ1:AM110']
    }

    # Create named styles
    date_style = NamedStyle(name='date_style', number_format='YYYY-MM-DD')
    no_fill = openpyxl.styles.PatternFill(fill_type=None)
    no_side = openpyxl.styles.Side(border_style=None)
    side = openpyxl.styles.Side(border_style='thin')
    no_border = openpyxl.styles.borders.Border(left=no_side, right=no_side, top=no_side, bottom=no_side)
    all_border = openpyxl.styles.borders.Border(left=side, right=side, top=side, bottom=side)

    wb = openpyxl.load_workbook(write_to_file)

    # add the styles
    wb.add_named_style(date_style)

    ws = wb['Monetization']

    for _pos in ['AF5', 'AG5', 'AK4', 'AL4']:
        ws[_pos].style = 'date_style'
    for _pos in ['AF4', 'AG4', 'AH4', 'AI4', 'AK3', 'AL3', 'AM3', 'AO4', 'AP4', 'AQ4', 'AX4', 'BM2', 'BP3']:
        ws[_pos].number_format = '0.0%'
    for _row in np.arange(8, 109):
        ws['BK{}'.format(_row)].number_format = '0.0%'
        ws['BT{}'.format(_row)].number_format = '0.0%'
    for _row in np.arange(7, 109):
        ws['BQ{}'.format(_row)].number_format = 'kr #, ##0'
        ws['BR{}'.format(_row)].number_format = 'kr #, ##0'
        ws['BS{}'.format(_row)].number_format = 'kr #, ##0'

    for color in list(colors.keys()):
        for my_pos in colors[color]:
            for my_row in ws[my_pos]:
                for my_cell in my_row:
                    my_cell.fill = PatternFill(patternType='solid', fgColor=color)

    for my_cell_pos in ['A5', 'B5', 'F1', 'F2', 'AF3', 'AG3', 'AI3', 'AF5', 'AF8', 'AG6', 'AN4', 'AO4', 'AP4',
                        'AQ4', 'AT4', 'AX4', 'AY4', 'BH8', 'BI8', 'BP3']:
        ws[my_cell_pos].fill = no_fill
        ws[my_cell_pos].border = all_border

    for rng in ['C8:C50', 'P8:P50', 'H4:W5', 'BH9:BI50', 'BO6:BT50', 'BO2:BO2', 'BQ3:BR3']:
        for my_row in ws[rng]:
            for my_cell in my_row:
                my_cell.border = all_border

    for rng in ['BP7:BR7']:
        for my_row in ws[rng]:
            for my_cell in my_row:
                my_cell.font = Font(bold=True)

    for rng in ['H1:O1', 'P1:Q1', 'R1:W1', 'Y1:AB1', 'AC1:AD1',
                'H4:K4', 'L4:O4', 'BG4:BL4', 'BO2:BT2', 'BQ3:BQ5', 'BR3:BT5']:
        ws.merge_cells(rng)
        ws[rng.split(':')[0]].alignment = Alignment(horizontal='center', vertical='center')
        ws[rng.split(':')[0]].font = Font(bold=True)

    wb.save(write_to_file)


def qc_plots(monetization_file, project_tag, plot_dir=None):
    """
    QC plots of a monetization file. The file must have been opened and saved in Excel first, so that the
    formulas have calculated values. Only the overview plot is currently created.
    """
    from Fossagrim.utils.definitions import standard_colors as scrs
    from Fossagrim.utils.definitions import standard_linestyles as sls
    from Fossagrim.utils.definitions import standard_linewidths as slw

    wb = openpyxl.load_workbook(monetization_file, data_only=True)
    ws = wb['Monetization']
    annual_climate_benefit_unit = ws['AC4'].value
    accum_climate_benefit_unit = ws['AD4'].value
    resampled_climate_benefit_unit = ws['AA4'].value
    overview_unit = ws['H4'].value
    contract_length = ws['AN4'].value
    wb.close()
    table = pd.read_excel(monetization_file, sheet_name='Monetization', header=5)
    if plot_dir is None:
        qc_plot_dir = os.path.join(os.path.split(monetization_file)[0], 'QC_plots')
    else:
        qc_plot_dir = plot_dir

    def annual_climate_benefit():
        x = table['t.1'].values
        y1 = table['30 yr contract'].values
        y2 = table['Unnamed: 28'].values
        y3 = table['100 yr contract'].values
        fig, ax = plt.subplots()
        ax.plot(x, y2, c=scrs['climate_benefit_0'],
                linestyle=sls['climate_benefit_0'],
                linewidth=slw['climate_benefit_0'],
                label='Climate benefit')
        ax.plot(x, y1, 'b-', label='30 yr contract')
        ax.plot(x, y3, 'b.', label='100 yr contract')
        ax.set_title('{} Annual Climate Benefit'.format(project_tag))
        ax.set_xlabel('Years')
        ax.set_ylabel(annual_climate_benefit_unit)
        ax.legend()
        ax.grid(True)
        fig.savefig(os.path.join(qc_plot_dir, '{} annual_climate_benefit.png'.format(project_tag)))

    def accumulated_climate_benefit():
        x = table['t.1'].values
        y1 = table['30 yr contract.1'].values
        y2 = table['Unnamed: 29'].values
        y3 = table['100 yr contract.1'].values
        fig, ax = plt.subplots()
        ax.plot(x, y2, c=scrs['climate_benefit_0'],
                linestyle=sls['climate_benefit_0'],
                linewidth=slw['climate_benefit_0'],
                label='Climate benefit')
        ax.plot(x, y1, 'b-', label='30 yr contract')
        ax.plot(x, y3, 'b.', label='100 yr contract')
        ax.set_title('{} Accumulated Climate Benefit'.format(project_tag))
        ax.set_xlabel('Years')
        ax.set_ylabel(accum_climate_benefit_unit)
        ax.legend()
        ax.grid(True)
        fig.savefig(os.path.join(qc_plot_dir, '{} accumulated_climate_benefit.png'.format(project_tag)))

    def resampled_climate_benefit():
        x = table['t.1'].values
        y1 = table['Unnamed: 27'].values
        y2 = table['Unnamed: 26'].values
        fig, ax = plt.subplots()
        ax.plot(x, y2, 'y.', label='Linear intpol / 5yr')
        ax.plot(x, y1, 'b-', label='Running average')
        ax.set_title('{} Resampled Climate Benefit'.format(project_tag))
        ax.set_xlabel('Years')
        ax.set_ylabel(resampled_climate_benefit_unit)
        ax.legend()
        ax.grid(True)
        fig.savefig(os.path.join(qc_plot_dir, '{} resampled_climate_benefit.png'.format(project_tag)))

    def overview():
        x = table['year'].values
        y1 = table['COPY OVER!.1'].values  # Base case
        y2 = table['COPY OVER!.2'].values  # Project case
        _y3 = table['Unnamed: 8'].values  # Product
        _y4 = table['Unnamed: 9'].values  # Substitution
        # shift the Product and Substitution arrays, and skip the last element
        y3 = np.insert(np.array([0.]), 1, _y3[:-1], axis=0)
        y4 = np.insert(np.array([0.]), 1, _y4[:-1], axis=0)
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(x, y1, c=scrs['bau_1'], label='Base case forest')
        ax.fill_between(x, y1, color=scrs['bau_1'])
        ax.plot(x, y1 + y3, c=scrs['products_1'], label='Base case product')
        ax.fill_between(x, y1 + y3, y1, color=scrs['products_1'])
        ax.plot(x, y1 + y4, c=scrs['substitution_0'], label='Base case substitution')
        ax.fill_between(x, y1 + y4, y1 + y3, color=scrs['substitution_0'])
        ax.plot(x, y2, c=scrs['project_case_0'], label='Project case')
        ax.set_title('{} Overview'.format(project_tag))
        ax.set_xlabel('Year')
        ax.set_ylabel(overview_unit)
        ax.set_xlim(np.nanmin(x) - 5, np.nanmin(x) + 155)
        ax.legend()
        ax.grid(True)
        fig.savefig(os.path.join(qc_plot_dir, '{} overview.png'.format(project_tag)))

    def farmed_offsets():
        x = table['t.1'].values
        y1 = table['Net farmed offsets / yr'].values
        y2 = table['Buffer reserved'].values
        y3 = table['Buffer released'].values
        y4 = table['Farmed offsets'].values
        fig, ax = plt.subplots(figsize=(10, 5))
        data = {
            'Net farmed offsets': y1,
            'Buffer reserved': y2,
            'Buffer released': y3
        }
        bottom = np.zeros(len(x))
        for label, height in data.items():
            ax.bar(x, height, label=label, bottom=bottom)
            bottom += height
        ax.plot(x, y4, label='Farmed offsets')
        ax.set_title('{} Farmed offsets'.format(project_tag))
        ax.set_ylabel('Ton')
        ax.set_xlabel('Time [years]')
        ax.legend()
        ax.grid(True)
        ax.set_xlim(0, contract_length)
        fig.savefig(os.path.join(qc_plot_dir, '{} farmed_offsets.png'.format(project_tag)))

    # annual_climate_benefit()
    # accumulated_climate_benefit()
    # resampled_climate_benefit()
    overview()
    # farmed_offsets()
