import os

import numpy as np
import pandas as pd
import pytest

from Fossagrim.io.common import (my_float, my_str, my_str_to_float, get_row_index, read_excel, write_csv_file,
                                 read_csv_file)
from Fossagrim.utils.definitions import heureka_standdata_keys, heureka_standdata_desc
from Fossagrim.unit_tests.conftest import UNIT_TEST_DIR


@pytest.mark.parametrize('value, expected', [('0,12', 0.12), ('0.12', 0.12), (3, 3.0), (1.5, 1.5), ('-2,5', -2.5)])
def test_my_float(value, expected):
    assert my_float(value) == pytest.approx(expected)


def test_my_float_rejects_text():
    with pytest.raises(ValueError):
        my_float('abc')


def test_my_str():
    assert my_str(None) == ''
    assert my_str(1.5) == '1.5'


@pytest.mark.parametrize('value, expected', [('3', 3), ('3.0', 3), ('3.5', 3.5), ('abc', 'abc'), ('', ''),
                                             (None, None)])
def test_my_str_to_float(value, expected):
    result = my_str_to_float(value)
    assert result == expected
    assert type(result) is type(expected)


def test_get_row_index_skips_empty_cells_and_supports_exact_match():
    table = pd.DataFrame({'Variable': [np.nan, 'Treatments:Year', 'Year', 'Soil Carbon Stock']})
    assert get_row_index(table, 'Variable', 'Year') == 1
    assert get_row_index(table, 'Variable', 'Year', exact=True) == 2
    assert get_row_index(table, 'Variable', 'Soil') == 3
    assert get_row_index(table, 'Variable', 'Missing') is None


def test_read_excel_handles_decimal_comma_file():
    """TestDecimalSigns.xlsx has the same number written with ',' and '.', formatted as general, text and number"""
    table = read_excel(os.path.join(UNIT_TEST_DIR, 'TestDecimalSigns.xlsx'), 1, 'Sheet1')
    assert list(table.keys()) == ['General Format', 'Text Format', 'Number Format']
    for key in table.keys():
        for value in table[key]:
            assert my_float(value) == pytest.approx(0.12)


def test_read_excel_strips_column_names(forvaltningsplan_file):
    table = read_excel(forvaltningsplan_file, 7, 'data')
    assert 'Gran' in table.keys()
    assert 'Teig' in table.keys()
    assert 'Vol\n/\ndaa' in table.keys()
    assert 'Bonitering\ntreslag' in table.keys()


def test_write_and_read_csv_file_round_trip(tmp_path, capsys):
    csv_file = str(tmp_path / 'stands.csv')
    write_csv_file(csv_file, heureka_standdata_keys, heureka_standdata_desc,
                   StandId='A', DGV=46, SiteIndexSpecies='G', V=np.float64(123.456), unknown_key=8)
    assert 'unknown_key' in capsys.readouterr().out
    write_csv_file(csv_file, heureka_standdata_keys, heureka_standdata_desc, append=True,
                   StandId='B', SiteIndexSpecies='T', Note='Åre æøå')

    with open(csv_file, encoding='cp1252') as f:
        lines = f.readlines()
    assert len(lines) == 4
    assert lines[1].rstrip('\n').split(';') == heureka_standdata_keys

    data, descriptors = read_csv_file(csv_file)
    assert descriptors == heureka_standdata_desc
    assert list(data.keys()) == heureka_standdata_keys  # no line break left in the last key
    assert data['StandId'] == ['A', 'B']
    assert data['DGV'] == [46, '']
    assert data['V'] == [123.46, '']  # floats are written with two decimals
    assert data['Note'] == ['', 'Åre æøå']
    assert data['SetAsideType'] == ['', '']


def test_write_csv_file_header_only(tmp_path):
    csv_file = str(tmp_path / 'header.csv')
    write_csv_file(csv_file, heureka_standdata_keys, heureka_standdata_desc)
    data, _ = read_csv_file(csv_file)
    assert all(len(values) == 0 for values in data.values())


def test_write_csv_file_requires_matching_descriptions(tmp_path):
    with pytest.raises(IOError):
        write_csv_file(str(tmp_path / 'x.csv'), ['a', 'b'], ['only one'])
