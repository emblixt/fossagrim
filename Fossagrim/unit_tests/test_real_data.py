"""
Tests against the real FHF23-0007 files in Fossagrim/test_data: the Forvaltningsplan, and the Heureka results
simulated from it. They are skipped automatically when the files are not available.
"""
import os

import numpy as np
import openpyxl
import pytest

from Fossagrim.io.common import read_excel, read_csv_file
from Fossagrim.io.stands import get_active_forests_from_stand, average_over_stands
from Fossagrim.utils.definitions import fossagrim_standdata_keys
from Fossagrim.unit_tests.conftest import REAL_PROJECT_NAME

ID = 'Fossagrim ID'


def test_real_file_has_all_expected_columns(real_forvaltningsplan):
    table = read_excel(real_forvaltningsplan, 7, 'data')
    missing = [key for key in fossagrim_standdata_keys if key not in table.keys()]
    assert missing == []


def test_real_active_forests(real_forvaltningsplan):
    active, areas, fractions = get_active_forests_from_stand(real_forvaltningsplan, stand_id_key=ID,
                                                             sheet_name='data')
    assert set(active.keys()) == {'Spruce', 'Pine', 'Birch'}
    assert sum(fractions.values()) == pytest.approx(1.)
    assert sum(len(stands) for stands in active.values()) == 162
    # Values printed by earlier runs of the scripts on this file
    assert sum(areas['Spruce']) == pytest.approx(1099.6)
    assert fractions['Pine'] == pytest.approx(0.648, abs=1e-3)


def test_real_average_over_stands(real_forvaltningsplan):
    active, areas, _ = get_active_forests_from_stand(real_forvaltningsplan, stand_id_key=ID, sheet_name='data')
    table = read_excel(real_forvaltningsplan, 7, 'data')
    avg = average_over_stands(active, table, ID, 'FHF23-0007 ')
    assert list(avg[ID]) == ['FHF23-0007 ' + key for key in active]
    for i, key in enumerate(active):
        assert avg['Prod.areal'][i] == pytest.approx(sum(areas[key]))


def test_real_project_settings(real_forvaltningsplan, tmp_path):
    from Fossagrim.create_heureka_input_files import project_settings

    _, _, average_over, _, result_file, result_sheets, combine_sheets, csv_file = \
        project_settings('FHF23-0007 v01', str(tmp_path), real_forvaltningsplan, create_output=True)

    assert openpyxl.load_workbook(result_file).sheetnames == result_sheets
    assert len(result_sheets) == 6
    assert len(combine_sheets) == 2
    data, _ = read_csv_file(csv_file)
    assert data['StandId'] == ['FHF23-0007 v01 Spruce', 'FHF23-0007 v01 Pine', 'FHF23-0007 v01 Birch']
    assert data['SiteIndexSpecies'] == ['G', 'T', 'B']
    for key in ['PropPine', 'PropSpruce', 'PropBirch']:
        assert all(0 <= value <= 1 for value in data[key])
    assert os.path.isfile(csv_file)


# ---------------------------------------------------------------------------------------------------------------
# Real Heureka results (FHF23-0007 v01_1), simulated from the average stands of the Forvaltningsplan above
# ---------------------------------------------------------------------------------------------------------------

def test_real_read_raw_heureka_results(real_heureka_results):
    from Fossagrim.io.heureka_results import read_raw_heureka_results, TOTAL_CARBON

    # The BAU sheets start with sub-periods (years 0, 1, 2) for FinalFelling, SoilPreparation and Planting
    bau = read_raw_heureka_results(real_heureka_results, REAL_PROJECT_NAME + ' Spruce BAU')
    years = [float(year) for year in bau['Year']]
    assert years[:4] == [0., 5., 10., 15.]
    assert all(year % 5 == 0 for year in years)
    assert len(years) == 31
    assert bau[TOTAL_CARBON].iloc[0] == pytest.approx(158.153)
    assert bau[TOTAL_CARBON].iloc[1] == pytest.approx(101.832)
    assert bau.attrs[TOTAL_CARBON] == 'ton C/ha'
    assert bau['Treatment'].iloc[0] == 'FinalFelling'

    pres = read_raw_heureka_results(real_heureka_results, REAL_PROJECT_NAME + ' Spruce PRES')
    assert [float(year) for year in pres['Year']] == years


def test_real_sheet_names_match_project_settings(real_forvaltningsplan, real_heureka_results, tmp_path):
    """The sheet names created by create_heureka_input_files.py must be the ones used in the real results file"""
    from Fossagrim.create_heureka_input_files import project_settings

    _, _, _, _, _, result_sheets, _, _ = project_settings(
        REAL_PROJECT_NAME, str(tmp_path), real_forvaltningsplan, create_output=False)
    assert openpyxl.load_workbook(real_heureka_results).sheetnames == result_sheets


def test_real_rearrange_heureka_results(real_forvaltningsplan, real_heureka_results, tmp_path):
    from Fossagrim.create_heureka_input_files import project_settings
    from Fossagrim.io.heureka_results import (read_raw_heureka_results, rearrange_raw_heureka_results,
                                              read_rearranged_heureka_results, TOTAL_CARBON)
    from Fossagrim.utils.monetization_parameters import variables_used_in_monetization

    _, _, _, _, _, result_sheets, combine_sheets, _ = project_settings(
        REAL_PROJECT_NAME, str(tmp_path), real_forvaltningsplan, create_output=False)
    rearrange_raw_heureka_results(real_heureka_results, result_sheets, combine_sheets)

    result = read_rearranged_heureka_results(real_heureka_results)
    assert list(result.keys()) == result_sheets + list(combine_sheets.keys())
    for table in result.values():
        assert len(table) == 31
        # all monetization variables except the UniqueID, which is not in this (older) result file
        assert set(table.keys()) == set(variables_used_in_monetization) - {'UserDefined8 : UniqueID'}

    # The rearranged tables contain the same values as the raw sheets
    raw = read_raw_heureka_results(real_heureka_results, REAL_PROJECT_NAME + ' Pine PRES')
    rearranged = result[REAL_PROJECT_NAME + ' Pine PRES']
    assert [float(x) for x in rearranged[TOTAL_CARBON]] == pytest.approx([float(x) for x in raw[TOTAL_CARBON]])

    # The combined tables are the area weighted sums of the three species
    for combined_name, sheets_and_fractions in combine_sheets.items():
        sheets, fractions = sheets_and_fractions[0::2], sheets_and_fractions[1::2]
        assert sum(fractions) == pytest.approx(1.)
        expected = sum(fraction * np.array([float(x) for x in result[sheet][TOTAL_CARBON]])
                       for sheet, fraction in zip(sheets, fractions))
        combined = result[combined_name]
        assert [float(x) for x in combined[TOTAL_CARBON]] == pytest.approx(list(expected))
        assert [float(x) for x in combined['Year']] == [float(x) for x in result[sheets[0]]['Year']]


def test_real_nature_and_climate_effect(real_heureka_results):
    from Fossagrim.io.heureka_results import get_nature_and_climate_effect

    # Regression values: 30 year sums of PRES - BAU, as calculated when this test was written
    carbon, dead_wood, recreation = get_nature_and_climate_effect(real_heureka_results, REAL_PROJECT_NAME + ' Spruce')
    assert carbon == pytest.approx(469.904)
    assert dead_wood == pytest.approx(56.572)
    assert recreation == pytest.approx(2.863)

    for species in ['Pine', 'Birch']:
        effects = get_nature_and_climate_effect(real_heureka_results, '{} {}'.format(REAL_PROJECT_NAME, species))
        assert None not in effects


# ---------------------------------------------------------------------------------------------------------------
# create_heureka_input_files.py: the average stands are correct area weighted averages of the active stands.
# The expected values are calculated directly from the Excel file with pandas, without any Fossagrim functions,
# so that a bug in the averaging or unit conversion can not hide itself.
# ---------------------------------------------------------------------------------------------------------------

DAA_PER_HA = 10.  # 1 ha = 10 daa, so a value per daa is multiplied by 10 to get the value per ha
M3FUB_TO_M3SK = 1.19  # Forvaltningsplan volumes are m3fub (under bark), Heureka volumes are m3sk (standing, over bark)
CSV_TOLERANCE = 0.006  # the csv file stores floats with two decimals
VOLUME_DESCRIPTION = 'Volym (m3sk/ha. produktiv areal)'  # description (line 1) of the volume column in the csv file
FOREST_TYPES = {'Spruce': 'Gran', 'Pine': 'Furu', 'Birch': 'Bjørk'}  # name in output: name in Forvaltningsplan


@pytest.fixture(scope='module')
def real_averaged_stand_csv(tmp_path_factory):
    """Runs create_heureka_input_files.py on the real Forvaltningsplan, and returns the csv file it creates"""
    from Fossagrim.unit_tests.conftest import REAL_FORVALTNINGSPLAN, run_script

    if not os.path.isfile(REAL_FORVALTNINGSPLAN):
        pytest.skip('Real test data {} is not available'.format(os.path.basename(REAL_FORVALTNINGSPLAN)))
    folder = str(tmp_path_factory.mktemp('FHF23-0007'))
    created = run_script('create_heureka_input_files.py', 'FHF23-0007 v01', folder, REAL_FORVALTNINGSPLAN,
                         cwd=folder)
    assert created.returncode == 0, created.stderr
    return os.path.join(folder, 'FHF23-0007 v01 Averaged stand data.csv')


@pytest.mark.parametrize('forest_type', list(FOREST_TYPES.keys()))
def test_real_average_volume_is_area_weighted(real_averaged_stand_csv, forest_type):
    import pandas as pd
    from Fossagrim.unit_tests.conftest import REAL_FORVALTNINGSPLAN

    # Volume of the average stand in the csv file, found through the column description in line 1
    with open(real_averaged_stand_csv, encoding='cp1252') as f:
        descriptions = f.readline().rstrip('\n').split(';')
    output = pd.read_csv(real_averaged_stand_csv, sep=';', header=1, encoding='cp1252')
    volume_column = output.columns[descriptions.index(VOLUME_DESCRIPTION)]
    row = output[output['StandId'] == 'FHF23-0007 v01 {}'.format(forest_type)]
    assert len(row) == 1, 'Average stand for {} not found in the csv file'.format(forest_type)
    csv_volume = float(row[volume_column].iloc[0])

    # Area weighted average of 'Vol / daa' over the active stands of this forest type in the Forvaltningsplan
    stands = pd.read_excel(REAL_FORVALTNINGSPLAN, sheet_name='data', header=7)
    stands.columns = [c.strip() if isinstance(c, str) else c for c in stands.columns]
    active = stands[(stands['Volum status'] == 1) &
                    (stands['Bonitering\ntreslag'] == FOREST_TYPES[forest_type])]
    assert len(active) > 0, 'No active {} stands in the Forvaltningsplan'.format(FOREST_TYPES[forest_type])
    area = active['Prod.areal']
    weighted_volume_per_daa = (area * active['Vol\n/\ndaa']).sum() / area.sum()  # m3fub/daa
    expected = weighted_volume_per_daa * DAA_PER_HA * M3FUB_TO_M3SK  # m3sk/ha

    assert csv_volume == pytest.approx(expected, abs=CSV_TOLERANCE), (
        '{}: {} active stands, area weighted volume {:.4f} m3fub/daa -> expected {:.3f} m3sk/ha, '
        'but the csv file has {:.3f}'.format(forest_type, len(active), weighted_volume_per_daa, expected, csv_volume))


# ---------------------------------------------------------------------------------------------------------------
# rearrange_heureka_results.py: the combined results are the area weighted sums of the forest types.
# The expected values are calculated directly from the Excel files with pandas/openpyxl, without any Fossagrim
# functions:
#   fraction of a forest type = sum of 'Prod.areal' of its active stands / sum of 'Prod.areal' of all active stands
#   combined value            = sum over forest types of (fraction x value in the raw Heureka sheet)
# for every variable and every 5 year period (Heureka sub-periods, years that are not a multiple of 5, are dropped)
# ---------------------------------------------------------------------------------------------------------------

COMBINED_FOREST_TYPES = ['Spruce', 'Pine', 'Birch']  # order used in the sheet names of the combined results
COPIED_VARIABLES = ['Treatment', 'Year']  # copied, not area weighted, in the combined results


@pytest.fixture(scope='module')
def real_rearranged_results(tmp_path_factory):
    """Runs rearrange_heureka_results.py on a copy of the real Heureka results, and returns the copy"""
    import shutil
    from Fossagrim.unit_tests.conftest import (REAL_FORVALTNINGSPLAN, REAL_HEUREKA_RESULTS, REAL_PROJECT_NAME,
                                               run_script)

    for f in [REAL_FORVALTNINGSPLAN, REAL_HEUREKA_RESULTS]:
        if not os.path.isfile(f):
            pytest.skip('Real test data {} is not available'.format(os.path.basename(f)))
    folder = str(tmp_path_factory.mktemp('FHF23-0007 rearrange'))
    results_file = os.path.join(folder, os.path.basename(REAL_HEUREKA_RESULTS))
    shutil.copy(REAL_HEUREKA_RESULTS, results_file)
    rearranged = run_script('rearrange_heureka_results.py', REAL_PROJECT_NAME, folder, REAL_FORVALTNINGSPLAN,
                            results_file, cwd=folder)
    assert rearranged.returncode == 0, rearranged.stderr
    return results_file


def independent_area_fractions():
    """Fraction of the total active productive area of each forest type, straight from the Forvaltningsplan"""
    import pandas as pd
    from Fossagrim.unit_tests.conftest import REAL_FORVALTNINGSPLAN

    stands = pd.read_excel(REAL_FORVALTNINGSPLAN, sheet_name='data', header=7)
    stands.columns = [c.strip() if isinstance(c, str) else c for c in stands.columns]
    active = stands[stands['Volum status'] == 1]
    total_area = active['Prod.areal'].sum()
    return {forest_type: active.loc[active['Bonitering\ntreslag'] == norwegian, 'Prod.areal'].sum() / total_area
            for forest_type, norwegian in FOREST_TYPES.items()}


def independent_raw_sheet(workbook, sheet_name):
    """
    The values of a raw Heureka sheet (Result type | Variable | Unit | Period ...) at the 5 year periods,
    as {variable: [value per period]}, read cell by cell with openpyxl
    """
    rows = {row[1]: list(row[3:]) for row in workbook[sheet_name].iter_rows(min_row=2, values_only=True)}
    keep = [i for i, year in enumerate(rows['Year']) if year is not None and float(year) % 5 == 0]
    return {variable: [values[i] for i in keep] for variable, values in rows.items()}


def rearranged_table(workbook, title):
    """
    One table in the 'Rearranged results' sheet: the title is in row 1, the variable names in row 3 (after the
    index column) and the values from row 4. Returns {variable: [value per period]}
    """
    ws = workbook['Rearranged results']
    start = [c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value == title]
    assert len(start) == 1, 'Table "{}" not found in the Rearranged results sheet'.format(title)
    table = {}
    col = start[0] + 1  # skip the index column
    while ws.cell(3, col).value is not None:
        table[ws.cell(3, col).value] = [ws.cell(r, col).value for r in range(4, ws.max_row + 1)]
        col += 1
    return table


def test_real_area_fractions_sum_to_one():
    from Fossagrim.unit_tests.conftest import REAL_FORVALTNINGSPLAN

    if not os.path.isfile(REAL_FORVALTNINGSPLAN):
        pytest.skip('Real test data is not available')
    fractions = independent_area_fractions()
    assert all(fraction > 0 for fraction in fractions.values())
    assert sum(fractions.values()) == pytest.approx(1.)


@pytest.mark.parametrize('method', ['BAU', 'PRES'])
def test_real_combined_results_are_area_weighted(real_rearranged_results, method):
    from Fossagrim.unit_tests.conftest import REAL_HEUREKA_RESULTS, REAL_PROJECT_NAME

    fractions = independent_area_fractions()
    raw_workbook = openpyxl.load_workbook(REAL_HEUREKA_RESULTS)  # the original, untouched, results
    raw = {forest_type: independent_raw_sheet(raw_workbook, '{} {} {}'.format(REAL_PROJECT_NAME, forest_type, method))
           for forest_type in COMBINED_FOREST_TYPES}

    combined_title = '{} Combined Stands {} {}'.format(REAL_PROJECT_NAME, '-and-'.join(COMBINED_FOREST_TYPES), method)
    combined = rearranged_table(openpyxl.load_workbook(real_rearranged_results), combined_title)

    # The combined table must contain the important carbon variables, so the test can not pass on an empty table
    for variable in ['Total Carbon Stock (dead wood, soil, trees, stumps and roots)', 'Soil Carbon Stock',
                     'Total Extracted Volume Fub (m³fub)', 'Year']:
        assert variable in combined, '{} is missing in the combined results'.format(variable)

    n_periods = len(raw['Spruce']['Year'])
    assert all(len(values) == n_periods for values in combined.values())

    # Year and Treatment are copied from the first forest type
    assert [float(x) for x in combined['Year']] == [float(x) for x in raw['Spruce']['Year']]
    assert [x if x not in [None, 'None'] else None for x in combined['Treatment']] == \
        [x if x not in [None, 'None'] else None for x in raw['Spruce']['Treatment']]

    for variable, values in combined.items():
        if variable in COPIED_VARIABLES:
            continue
        expected = [sum(fractions[forest_type] * float(raw[forest_type][variable][period])
                        for forest_type in COMBINED_FOREST_TYPES)
                    for period in range(n_periods)]
        assert [float(x) for x in values] == pytest.approx(expected, rel=1e-9), \
            '{} {}: combined values differ from the area weighted sum of {}'.format(
                method, variable, ', '.join(COMBINED_FOREST_TYPES))
