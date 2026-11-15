import numpy as np
import pandas as pd
import pytest

from Fossagrim.qc_input_pdfs import number_after, check_forvaltningsplan_table


def test_number_after():
    lines = ['Tømmervolum', '12 345,6', 'Antall bestand', '17']
    assert number_after(lines, 'Tømmervolum') == pytest.approx(12345.6)
    assert number_after(lines, 'Antall bestand') == 17
    with pytest.raises(ValueError):
        number_after(lines, 'Missing')


def forvaltning_table(**overrides):
    """A minimal 'Forvaltning' sheet. The first row holds the totals, as in the real files"""
    data = {
        'GBTBestand': ['Total', 1, 2, 3, 4],
        'Miljøfig': [np.nan, 'Nei', 'Ja', 'Ja', np.nan],
        'H.kl': [0, 5, 5, 3, 4],
        'Proj areal': [30., 10., 5., 0., 15.],
        'Proj vol': [300., 100., 50., 20., 130.],
        'Begrunnelse': [np.nan, np.nan, np.nan, np.nan, np.nan],
        'Prod.areal': [40., 10., 5., 10., 15.],
    }
    data.update(overrides)
    return pd.DataFrame(data)


def test_check_forvaltningsplan_finds_problems():
    log_text, warn_text = check_forvaltningsplan_table(forvaltning_table(), '', '')
    assert 'Bestand 2 has a mismatch between "misfig"' in warn_text  # MIS figure in the project area
    assert 'Bestand 3' not in warn_text.split('misfig')[0]  # MIS figure outside the project area is fine
    assert 'Bestand 3 is cutting in hogstklasse < 4' in warn_text
    assert 'Bestand 4 is cutting in hogstklasse 4 (4) without "Begrunnelse"' in warn_text
    assert '1 number of mis figs' in log_text


def test_check_forvaltningsplan_accepts_justified_harvest():
    table = forvaltning_table(**{'Begrunnelse': [np.nan, np.nan, np.nan, np.nan, 'Stormskadet'],
                                 'H.kl': [0, 5, 5, 5, 4], 'Miljøfig': [np.nan] * 5})
    _, warn_text = check_forvaltningsplan_table(table, '', '')
    assert warn_text == ''


def test_check_forvaltningsplan_detects_numbers_as_text():
    table = forvaltning_table(**{'Prod.areal': [40., 10., '5,0', 10., 15.]})
    _, warn_text = check_forvaltningsplan_table(table, '', '')
    assert "Column Prod.areal in Forvaltningsplan likely contains a mixture" in warn_text


def test_check_forvaltningsplan_duplicated_columns():
    table = forvaltning_table()
    table.insert(len(table.columns), 'Gran', [10., 1., 2., 3., 4.], allow_duplicates=True)
    table.insert(len(table.columns), 'Gran', [10., 1., 'x', 3., 4.], allow_duplicates=True)
    _, warn_text = check_forvaltningsplan_table(table, '', '')
    assert 'Column Gran:2' in warn_text
    assert 'Column Gran:1' not in warn_text
