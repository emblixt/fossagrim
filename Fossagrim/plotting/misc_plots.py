import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import os
import Fossagrim.utils.projects as fup
import Fossagrim.io.fossagrim_io as fio

markers = ['o', 'v', '^', '<', '>', 's', 'p', 'P', '*', 'X', 'D']


def find_first_index(table, key, value):
    found = False
    i = None
    if isinstance(key, list) and isinstance(value, list):
        for i, this_value in enumerate(table[key[0]]):
            if (value[0] in this_value) and (value[1] in table[key[1]][i]):
                found = True
                break
    else:
        for i, this_value in enumerate(table[key]):
            if this_value == value:
                found = True
                break
    if found:
        return i
    else:
        return None


def plot_plant_density():
    projects = ['FHF23-0{}'.format(_x) for _x in ['03', '04', '05', '06', '07', '08', '09', '10', '12']]
    spruce_sis = [];
    pine_sis = [];
    spruce_plant_den = [];
    pine_plant_den = []

    for project_tag in projects:
        print(project_tag)
        _, _, _, _, _, _, _, csv_stand_file, csv_treatment_file = fup.project_settings(project_tag)
        this_table = pd.read_csv(csv_stand_file, sep=';', header=1, encoding="ISO-8859-1")
        i_spruce = find_first_index(this_table, 'SiteIndexSpecies', 'G')
        i_pine = find_first_index(this_table, 'SiteIndexSpecies', 'T')
        if i_spruce is None:
            spruce_sis.append(np.nan)
        else:
            spruce_sis.append(this_table['SIS'][i_spruce])

        if i_pine is None:
            pine_sis.append(np.nan)
        else:
            pine_sis.append(this_table['SIS'][i_pine])

        this_table = pd.read_csv(csv_treatment_file, sep=';', header=1, encoding="ISO-8859-1")
        i_spruce = find_first_index(this_table, ['StandId', 'Treatment'], ['Spruce', 'Planting'])
        i_pine = find_first_index(this_table, ['StandId', 'Treatment'], ['Pine', 'Planting'])
        if i_spruce is None:
            spruce_plant_den.append(np.nan)
        else:
            spruce_plant_den.append(this_table['PlantDensity'][i_spruce])
        if i_pine is None:
            pine_plant_den.append(np.nan)
        else:
            pine_plant_den.append(this_table['PlantDensity'][i_pine])

    fig, ax = plt.subplots()
    for i, project_tag in enumerate(projects):
        ax.scatter(spruce_sis[i], spruce_plant_den[i], c='b', marker=markers[i], label=project_tag)
        ax.scatter(pine_sis[i], pine_plant_den[i], c='g', marker=markers[i], label='_nolegend_')

    ax.set_xlabel('SIS')
    ax.set_ylabel('Plant density [plants/ha]')
    ax.set_title('Plant density vs Bonity (blue = Spruce)')
    ax.legend()
    plt.show()

    return spruce_sis, pine_sis, spruce_plant_den, pine_plant_den


def plot_raw_data(data_dict, data_tag, my_title, qc_plot_dir):
    fig, ax = plt.subplots(figsize=(12, 12))
    _x = data_dict['Year']
    _count = 0
    _lw = 1
    _ls = '-'
    for key, _y in data_dict.items():
        if np.mod(_count, 3) == 0:
            _lw = 2
            _ls = '--'
        elif np.mod(_count, 3) == 1:
            _lw = 1
            _ls = '-'
        elif np.mod(_count, 3) == 2:
            _lw = 2
            _ls = ':'
        if key in ['Year', 'Treatment', 'Unit']:
            continue
        ax.plot(_x, _y, linewidth=_lw, linestyle=_ls, label=key)
        _count += 1
    ax.set_title('{} {}'.format(my_title, data_tag))
    ax.set_xlabel('Year')
    ax.grid(True)
    ax.legend()
    fig.savefig(os.path.join(qc_plot_dir, '{}_{}.png'.format(my_title.lower().replace(' ', '_'), data_tag)))


def plot_from_heureka_results(result_file, sheets, params=None, ax=None, x_key='Year',
                              diff_sheets=False, barchart=False, year_zero=None,  **kwargs):
    """
    Flexible plot of data stored in typical "Heureka results.xlsx" files
    :param result_file:
        str
        full path name of the Heureka results Excel file
    :param sheets:
        str or list str
        each string is the name of a sheet in the above results Excel file
    :param params:
        str of list of str
        each string is the (partial) name of a parameter listed in the above results Excel file
    :param ax:
    :param x_key:
    :param diff_sheets:
        bool
        If True, it plots the difference sheet[0] - sheet[1] for each listed param.
        All other listed sheets are ignored
    :param year_zero:
        float
        When 'x_key' is Year, this value is added, so it acts as a start year
    :param kwargs:
        optional keywords passed on to plot function
    :return:
    """
    if ax is None:
        fig, ax = plt.subplots(figsize=(10, 10))
    if params is None:
        params = 'Total Carbon Stock'
    if isinstance(params, str):
        params = [params]
    if isinstance(sheets, str):
        sheets = [sheets]
    if diff_sheets and len(sheets) < 2:
        raise IOError('sheets must contain at least two sheet names')
    if year_zero is None:
        year_zero = 0.

    width = kwargs.pop('width', 4.0)

    # read the different sheets from the results Excel file
    data = []
    for sheet in sheets:
        data.append(fio.read_raw_heureka_results(result_file, sheet))

    # check if parameters are available in all sheets
    for d, sheet in zip(data, sheets):
        for p in params:
            param_not_found = True
            check = [p.lower() in _x.lower() for _x in list(d.keys())]
            if True in check:
                param_not_found = False
            if param_not_found:
                raise IOError('Parameter {} not found in sheet {}'.format(p, sheet))

    y_unit = ''
    y_previous = None
    last_sheet = None
    bottom_params = None
    bottom_sheets = None
    i = 0
    for d, sheet in zip(data, sheets):
        x = d[[_x for _x in list(d.keys()) if x_key.lower() in _x.lower()][0]].values
        if x_key == 'Year':
            x += year_zero
        if i == 0:
            bottom_sheets = np.zeros(len(x))
        for j, p in enumerate(params):
            y_key = [_x for _x in list(d.keys()) if p.lower() in _x.lower()][0]
            y = np.asarray(d[y_key].values, dtype=np.float32)
            y_unit = d.attrs[y_key]
            if diff_sheets and y_previous is None:
                y_previous = y
                last_sheet = sheet
            elif diff_sheets and y_previous is not None:
                print(y_previous[:2], y[:2])
                diff = y_previous - y
                ax.plot(x, diff, label='{} - {}'.format(last_sheet, sheet), **kwargs)
                break
            if not diff_sheets:
                if barchart:
                    if j == 0:
                        bottom_params = np.zeros(len(x))
                    ax.bar(x, y, label='{} {}'.format(p, sheet), width=width,
                           bottom=bottom_params + bottom_sheets, **kwargs)
                    bottom_params += y
                else:
                    ax.plot(x, y, label='{} {}'.format(p, sheet), **kwargs)
        bottom_sheets += bottom_params
        i += 1

    ax.set_title(os.path.basename(result_file))
    ax.set_xlabel(x_key)
    ax.set_ylabel(y_unit)
    ax.grid(True)
    ax.legend()
