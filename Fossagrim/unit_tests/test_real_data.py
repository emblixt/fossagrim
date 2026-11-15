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
