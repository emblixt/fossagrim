import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import os
import unittest

from bokeh.plotting import figure, show
from bokeh.layouts import row, column
from bokeh.models import (ColumnDataSource, Select,
                          CustomJS, CustomJSTransform, LinearColorMapper)
from bokeh.models import PanTool,WheelZoomTool, ResetTool, SaveTool, CrosshairTool, HoverTool, ColorBar, LogColorMapper
from bokeh.models import (DataTable, NumberFormatter, SelectEditor, StringEditor, StringFormatter,
                          IntEditor, TableColumn, CheckboxEditor, Div)
from bokeh.io import output_file

import Fossagrim.utils.projects as fup
import Fossagrim.io.fossagrim_io as fio

markers = ['o', 'v', '^', '<', '>', 's', 'p', 'P', '*', 'X', 'D']

output_file('C:\\Users\marte\Documents\plot.html')

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
                              diff_sheets=False, barchart=False, year_zero=None,
                              save_plot_to=None,  **kwargs):
    """
    Flexible plot of data stored in typical "Heureka results.xlsx" files
    :param result_file:
        str
        full path name of the Heureka results Excel file
    :param sheets:
        str or list str
        each string is the name of a sheet in the above results Excel file
    :param params:
        str or list of str
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
    :param save_plot_to:
        str
        full file name to which the plot is saved
    :param kwargs:
        optional keywords passed on to plot function
    :return:
        None by default
        if diff_sheets is True, it returns the difference
    """
    fig = None
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

    diff = None
    width = kwargs.pop('width', 4.0)

    # read the different sheets from the results Excel file
    data = []
    for sheet in sheets:
        try:
            data.append(fio.read_raw_heureka_results(result_file, sheet))
        except KeyError as error:
            print(error)
            continue

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
            bottom_sheets = np.zeros(len(x), dtype='float64')
        for j, p in enumerate(params):
            y_key = [_x for _x in list(d.keys()) if p.lower() in _x.lower()][0]
            y = np.asarray(d[y_key].values, dtype=np.float32)
            y_unit = d.attrs[y_key]
            if diff_sheets and y_previous is None:
                y_previous = y
                last_sheet = sheet
            elif diff_sheets and y_previous is not None:
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

        if not diff_sheets:
            raise NotImplementedError('Problem with bottom_sheets - need check')
        # TODO
        # XXX
        # problems with the line below, which is really necessary when diff_sheets is False
        # bottom_sheets += bottom_params
        i += 1

    ax.set_title(os.path.basename(result_file))
    ax.set_xlabel(x_key)
    ax.set_ylabel(y_unit)
    ax.grid(True)
    ax.legend()

    if save_plot_to is not None:
        if fig is None:
            fig = ax.get_figure()
        fig.savefig(save_plot_to)

    return diff


def qc_stand_data(stand_data_table, stand_id, qc_plot_dir):
    """
    Creates QC plots of the stand data to highlight problems with the input before we start
    :param stand_data_table
        panda table
        Output from fossagrim_io.read_excel
    """
    from Fossagrim.utils.definitions import fossagrim_standdata_keys

    text_style = {'fontsize': 'x-small', 'bbox': {'facecolor': 'w', 'alpha': 0.5}}
    text_style_flag = {'fontsize': 'x-small', 'bbox': {'facecolor': 'r', 'alpha': 0.5}}
    flierprops = dict(markerfacecolor = 'r')

    fig1, axs1 = plt.subplots(13, 1, figsize=(8, 12))
    fig1.subplots_adjust(hspace=0)
    fig1.suptitle(stand_id)
    fig2, axs2 = plt.subplots(13, 1, figsize=(8, 12))
    fig2.subplots_adjust(hspace=0)
    fig2.suptitle(stand_id)
    fig3, axs3 = plt.subplots(13, 1, figsize=(8, 12))
    fig3.subplots_adjust(hspace=0)
    fig3.suptitle(stand_id)
    #fig4, axs4 = plt.subplots(13, 1, figsize=(8, 12))
    #fig4.subplots_adjust(hspace=0)

    i = 0
    for _key in list(stand_data_table.keys())[:38]:
    # for _key in fossagrim_standdata_keys:
        # The number 38 should reflect the part of the stand data table which contains the necessary data
        # if _key not in [_x.strip() for _x in list(stand_data_table.keys())[:38]]:
        #    error_txt = 'ERROR! Necessary key {} is not in stand data table'.format(_key)
        #    print(error_txt)
        #    raise IOError(error_txt)
        # empty columns
        if 'Unnamed' in _key:
            continue
        # Columns that only should contain strings
        if _key in ['Miljøfig', 'Bonitering\ntreslag', 'Tetthet', 'Fossagrim ID']:
            continue
        if i < 13:
            axs = axs1
            ii = i
        elif i < 26:
            axs = axs2
            ii = i - 13
        else:
            axs = axs3
            ii = i - 26
        # else:
        #     axs = axs4
        #     ii = i - 39
        if stand_data_table[_key].values.dtype == 'object':
            if len(stand_data_table[_key][stand_data_table[_key].isna()]) > 0:
                axs[ii].plot([0, 1], lw=0)
                axs[ii].text(0.1, 0.9, _key, ha='left', va='top', transform=axs[ii].transAxes, **text_style_flag)
                axs[ii].text(0.5, 0.5, 'Contain NaNs', ha='left', va='top', transform=axs[ii].transAxes, **text_style_flag)
            else:
                print('WARNING: Key {} can not be plotted'.format(_key))
                # print(stand_data_table[_key])
                continue
        else:
            boxplot = axs[ii].boxplot(stand_data_table[_key].values, vert=False, flierprops=flierprops)
            if len(boxplot['fliers'][0].get_xdata()) < 1:
                axs[ii].text(0.1, 0.9, _key, ha='left', va='top', transform=axs[ii].transAxes, **text_style)
            else:
                axs[ii].text(0.1, 0.9, _key, ha='left', va='top', transform=axs[ii].transAxes, **text_style_flag)
        axs[ii].tick_params(axis="x", direction="in", pad=-12, labelsize=8)
        axs[ii].set_yticks([], [])

        i += 1

    for i, fig in enumerate([fig1, fig2, fig3]):
        fig.savefig(os.path.join(qc_plot_dir, 'bestand qc {}.png'.format(i+1)))


def list_duplicates(seq):
    seen = set()
    seen_add = seen.add
    # adds all elements it doesn't know yet to seen and all other to seen_twice
    seen_twice = set( x for x in seq if x in seen or seen_add(x) )
    # turn the set into a list (as requested)
    return list(seen_twice)


def plot_collected_stand_data(csv_file):
    """
    Uses Bokeh to create an interactive scatter plot of
    :param csv_file:
    :return:
    """
    info_text = """
    <h1> Information</h1>
    This plot shows some modelled climate and nature benefits that can be gained by invoking a 30 year
    preservation scheme on the displayed forests.
    The forests (StandID) are listed in the table at the bottom. Each forest has been modelled twice, once 
    as a preserved forest, and once as <i>business-as-usual</i> with clear cutting and planting.
    The effect is simply calculated by taking the difference (preservation - business-as-usual) and summing that
    difference over 30 year - which is the contract length suggested by Fossagrim.
    The one line table just beneath the plot shows the total effect from all the forests, where the Age is the
    areal weighted mean of the age of the forest at the beginning of the 30 year contract.
    
    <h2>Interactivity</h2>
    <ul>
     <li>You can zoom and pan in the plot using the toolbar on top right corner</li>
     <li>With the <i>Hover</i> tool active, you can see which forest each point comes from by hovering
        over it with your mouse</li>
     <li>You can decide which parameter should be plotted along the X- and Y-axis, and marker size, by 
        using the three dropdown menus right of the plot</li>
    </ul>
        
    <h2>Columns</h2>
    The columns of the lower table are:
    <ul>
     <li><b>Stand ID:</b> The unique name of the forest (stand) used by Fossagrim</li>
     <li><b>Area:</b> Productive area, in hectar, of the preserved forest</li>
     <li><b>Age:</b> Mean age of the trees at the beginning of the 30 year contract period</li>
     <li><b>Tree:</b> Dominant tree species in the forest</li>
     <li><b>Reduced emissions:</b> The difference in bounded CO2 by living and dead trees and roots. between 
        the <i>preservation</i> case and <i>business-as-usual</i> case. The difference is then summed over
        the 30 year contract length. This way of counting neglects the 
        effect from C02 stored in long-lived wood products and in the soil</li>
     <li><b>Increase in deadwood volume:</b>The difference in volume of deadwood (standing and lying down) 
       with a diameter larger than 20cm, 
       between the <i>preservation</i> case and <i>business-as-usual</i> case. The difference is then summed over
        the 30 year contract length. </li>
     <li><b>Recreation effect:</b> The difference in the <a href="https://www.heurekaslu.se/wiki/Recreation_Value_Model">recreation index</a> 
       between the <i>preservation</i> case and <i>business-as-usual</i> case. The difference is then summed over
        the 30 year contract length. </li>
     
    </ul>
     
    """
    tools = "pan,wheel_zoom, box_zoom, crosshair, hover, reset"

    param_dict = {
        'Area [ha]': 'ProdArea',
        'Age': 'MeanAge',
        'Climate effect [ton CO2/ha]': 'UserDefinedVariable5_ClimateEffect',
        'Deadwood [m3/ha]': 'UserDefinedVariable6_DeadWoodEffect',
        'Recreation': 'UserDefinedVariable7_RecreationEffect',
        'SIS': 'SIS'
    }
    # Create selectors
    x_drop = Select(title='X-axis', value='Age', options=list(param_dict.keys()))
    y_drop = Select(title='Y-axis', value='SIS', options=list(param_dict.keys()))
    s_drop = Select(title='Size', value='Area [ha]', options=list(param_dict.keys()))

    # read csv file and drop non-necessary columns
    table = pd.read_csv(csv_file, sep=';', header=1, encoding='ISO-8859-1')
    drop_these = ['TotalArea', 'CountyCode', 'Altitude', 'Latitude', 'SoilMoistureCode',
                'VegetationType', 'Peat', 'InventoryYear', 'DGV', 'DG', 'SoilMoistureCode', 'VegetationType',
                'Peat', 'InventoryYear', 'DGV', 'DG', 'G', 'V', 'PropPine', 'PropSpruce', 'PropBirch', 'PropAspen',
                'PropOak', 'PropBeech', 'PropSouthernBroadleaf', 'PropContorta', 'PropOtherBroadleaf', 'AreaLevel2',
                'AreaLevel3', 'ParentStandId', 'ImpArea', 'NCArea', 'Layer', 'Register', 'CoordEast', 'CoordNorth',
                'DistanceToCoast', 'RoadId1', 'RoadId2', 'ClimateCode', 'SKSManagementClass', 'MaturityClass',
                'BlueTargetClass', 'SIH', 'SI_Management', 'BottomLayer', 'SoilDepth', 'SoilWater', 'Texture',
                'SlopeDirectionNorthEast', 'Ditch', 'SoilBearingCapacity', 'Surface', 'SlopeType', 'TerrRoadSlope',
                'OwnerType', 'EvenAgedCode', 'CAI', 'DiameterType', 'DGPine', 'DGSpruce', 'DGBirch', 'DGAspen', 'DGOak',
                'DGBeech', 'DGSouthernBroadleaf', 'DGContorta', 'DGOtherBroadleaf', 'HPine', 'HSpruce', 'HBirch',
                'HAspen', 'HOak', 'HBeech', 'HSouthernBroadleaf', 'HContorta', 'HOtherBroadleaf', 'SpeciesUser',
                'ProportionSpeciesUser', 'MeanDiameterSpeciesUser', 'MeanHeightSpeciesUser', 'DeadWoodTotal',
                'DeadWoodDecayClass1', 'DeadWoodDecayClass2', 'DeadWoodDecayClass3', 'DeadWoodDecayClass4',
                'DeadWoodDecayClass5', 'TerrainTransportDistance', 'LastClearcutYear', 'LastThinningYear',
                'LastFertilizationYear', 'LastRegenerationYear', 'RegenerationMethod', 'RegenerationSpecies',
                'RegenerationBreeded', 'Note', 'UserDefinedVariable1_TaxType', 'UserDefinedVariable2_RotationPeriod',
                'UserDefinedVariable3_ThinningYear', 'UserDefinedVariable4_PlantDensity', 'UserDefinedVariable8',
                'UserDefinedVariable9', 'UserDefinedVariable10', 'SetAsideType']
    table = table.drop(columns=drop_these)

    # Replace the tree types with human readable text, and associate a color each tree type
    tt = table['SiteIndexSpecies'].values
    colors = ['green'] * len(tt)
    for i, _item in enumerate(tt):
        if _item == 'T':
            tt[i] = 'Pine'
        elif _item == 'G':
            tt[i] = 'Spruce'
            colors[i] = 'orange'
        elif _item == 'B':
            tt[i] = 'Birch'
            colors[i] = 'blue'
    table['SiteIndexSpecies'] = tt

    # Convert from ton C to ton C02  ## and take into account area
    _ton_c = table['UserDefinedVariable5_ClimateEffect'].values
    # table['UserDefinedVariable5_ClimateEffect'] = (_ton_c * 44 / 12) * table['ProdArea'].values
    table['UserDefinedVariable5_ClimateEffect'] = (_ton_c * 44 / 12)

    # # Calculate total volume of Deadwood
    # _vol = table['UserDefinedVariable6_DeadWoodEffect'].values
    # table['UserDefinedVariable6_DeadWoodEffect'] = _vol * table['ProdArea'].values

    # Create the table of Stands and its source
    table_columns = [
        TableColumn(field='StandId', title='Stand ID', formatter=StringFormatter(font_style='bold'),
                    width=180),
        TableColumn(field='ProdArea', title='Area [ha]', width=80),
        TableColumn(field='MeanAge', title='Age', width=60),
        TableColumn(field='SiteIndexSpecies', title='Tree', width=60),
        TableColumn(field='UserDefinedVariable5_ClimateEffect',
                    formatter=NumberFormatter(format=f'0.000'),
                    title='Reduced emissions [ton CO2/ha]'),
        TableColumn(field='UserDefinedVariable6_DeadWoodEffect', title='Increase in deadwood volume [m3/ha]',
                    width=360),
        TableColumn(field='UserDefinedVariable7_RecreationEffect', title='Recreation effect'),
        TableColumn(field='SIS', title='SIS', width=40)
    ]
    table_source = ColumnDataSource(table)
    stands_table = DataTable(source=table_source, columns=table_columns, width=1000)

    # Create a summary table of data from above table
    sum_table = dict(
        title=['Total effect:'],
        area=[np.nansum(table['ProdArea'].values)],
        age=[np.nansum(table['MeanAge'].values * table['ProdArea'].values) / np.nansum(table['ProdArea'].values)],
        climate_effect=[np.nansum(table['UserDefinedVariable5_ClimateEffect'].values * table['ProdArea'].values) * 1.E-6],
        deadwood_effect=[np.nansum(table['UserDefinedVariable6_DeadWoodEffect'].values * table['ProdArea'].values)],
        recreation_effect=[np.nansum(table['UserDefinedVariable7_RecreationEffect'].values)]
    )
    sum_table_columns = [
        TableColumn(field='title', title='', formatter=StringFormatter(font_style='bold'), width=160),
        TableColumn(field='area', title='Area [ha]', width=80),
        TableColumn(field='age',
                    formatter=NumberFormatter(format=f'0.0'),
                    title='Age', width=60),
        TableColumn(field='climate_effect',
                    formatter=NumberFormatter(format=f'0.000'),
                    title='Reduced emissions [M ton CO2]'),
        TableColumn(field='deadwood_effect',
                    formatter=NumberFormatter(format=f'0.000'),
                    title='Increase in deadwood volume [m3]',
                    width=300),
        TableColumn(field='recreation_effect',
                    formatter=NumberFormatter(format=f'0.000'),
                    title='Recreation effect',
                    width=150)
    ]
    sum_table_source = ColumnDataSource(sum_table)
    sum_stands_table = DataTable(source=sum_table_source, columns=sum_table_columns, height=80, width=800)

    param_min_max = {
        _key: [np.nanmin(table_source.data[param_dict[_key]]), np.nanmax(table_source.data[param_dict[_key]])] for _key in param_dict.keys()
    }

    initial_size = 10. + 70 * (table_source.data[param_dict[s_drop.value]] - param_min_max[s_drop.value][0]) / \
            (param_min_max[s_drop.value][1] - param_min_max[s_drop.value][0])

    data_source_dict = dict(
        x=table_source.data[param_dict[x_drop.value]],
        y=table_source.data[param_dict[y_drop.value]],
        size=initial_size,
        color=colors,
        groups=table_source.data['SiteIndexSpecies'],
        stands=table_source.data['StandId']
    )
    data_source = ColumnDataSource(data_source_dict)

    # exp_cmap = LinearColorMapper(palette="Viridis256",
    #                              low=min(table_source.data['MeanAge']),
    #                              high=max(table_source.data['MeanAge']))

    p = figure(width=800, height=600, tools=tools)
    p.toolbar.logo = None
    hover = p.select(dict(type=HoverTool))
    hover.tooltips = [("(x,y)", "($x, $y)"), ("Stand", "@stands")]
    # p.scatter(x='x', y='y', size='size', line_color = None, fill_color = {"field": "color", "transform": exp_cmap},
    #           source=data_source)
    p.scatter(x='x', y='y', size='size', line_color=None, color='color', legend_field='groups',
              alpha=0.7, source=data_source)
    # bar = ColorBar(color_mapper=exp_cmap, location=(0, 0))
    # p.add_layout(bar, "left")
    p.xaxis.axis_label = x_drop.value
    p.yaxis.axis_label = y_drop.value
    title_text = 'Nature and climate effects of a 30 year program, size proportional to {}'.format(
        s_drop.value)
    p.title = title_text
    # bar.title = 'Age'

    code = """
    const x_param = x_drop.value;
    const y_param = y_drop.value;
    const s_param = s_drop.value;
    const min = min_max[s_param][0];
    const max = min_max[s_param][1];
    const s = [];
    function size(arr1) {
        //let s = [];
        for (let i = 0; i < arr1.length; i++) {
          s.push(10 + 70 * (arr1[i] - min) / (max - min));
              }
        return s;
    }
    data.data['size'] = size(table_data.data[params[s_param]]);
    data.data['x'] = table_data.data[params[x_param]];
    data.data['y'] = table_data.data[params[y_param]];
    xaxis.axis_label = x_param;
    yaxis.axis_label = y_param;
    title.text = 'Nature and climate effects of a 30 year program, size proportional to ' + s_param
    //bar.title = 'YYY'
    console.log('dropdown: ' + cb_obj.value, x_param);
    data.change.emit()
    """
    args_dict  = dict(x_drop=x_drop,
                      y_drop=y_drop,
                      s_drop=s_drop,
                      data=data_source,
                      table_data=table_source,
                      params=param_dict,
                      min_max=param_min_max,
                      # cmap=exp_cmap,
                      title=p.title,
                      xaxis=p.xaxis[0],
                      yaxis=p.yaxis[0],
                      # bar=bar
                      )

    x_drop.js_on_change('value',
                        CustomJS( args=args_dict, code=code))

    y_drop.js_on_change('value',
                        CustomJS( args=args_dict, code=code))

    s_drop.js_on_change('value',
                        CustomJS( args=args_dict, code=code))

    div = Div(text=info_text, width=600)
    show(column(row(p, column(x_drop, y_drop, s_drop), div), sum_stands_table, stands_table))


class TestCases(unittest.TestCase):
    def test_list_duplicates(self):
        print(list_duplicates(['A', 'B', 'C', 'A']))
        print(list_duplicates([1, 2, 3, 1]))

    def test_qc_stand_data(self):
        f = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0014 Arne Tag\\Bestandsoversikt 16052024.xlsx"
        table = fio.read_excel(f, 7, 0)
        qc_stand_data(table, os.path.basename(f), 'C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-014 Arne Tag\\QC_plots')
        plt.show()

    def test_plot_heureka_results(self):
        result_file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\FHF24-0016 Margrete Folsland\\FHF24-0016 Heureka results.xlsx"
        sheets = ["FHF24-0016 Pine PRES", "FHF24-0016 Pine BAU"]
        params = "Total Carbon Stock (dead wood, soil, trees, stumps and roots)"
        diff_sheets = True
        diff = plot_from_heureka_results(
            result_file,
            sheets=sheets,
            params=params,
            diff_sheets=diff_sheets
        )
        print(diff)
        plt.show()

    def test_plot_collected_stand_data(self):
       csv_file = "C:\\Users\\marte\\OneDrive - Fossagrim AS\\Prosjektskoger\\TEST COLLECT STAND DATA.csv"
       plot_collected_stand_data(csv_file)
