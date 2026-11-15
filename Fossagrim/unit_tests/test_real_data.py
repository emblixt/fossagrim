"""
Tests against the real FHF23-0007 Forvaltningsplan in Fossagrim/test_data.
They are skipped automatically when the file is not available.
"""
import os

import openpyxl
import pytest

from Fossagrim.io.common import read_excel, read_csv_file
from Fossagrim.io.stands import get_active_forests_from_stand, average_over_stands
from Fossagrim.utils.definitions import fossagrim_standdata_keys

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
