import openpyxl
import pandas as pd
import pytest

from Fossagrim.io.heureka_results import (read_raw_heureka_results, rearrange_raw_heureka_results,
                                          read_rearranged_heureka_results, combine_tables,
                                          combine_raw_heureka_results, get_nature_and_climate_effect,
                                          TOTAL_CARBON, SOIL_CARBON)
from Fossagrim.unit_tests.conftest import raw_heureka_rows

# 2027 is a sub-period (not a multiple of 5) and must be dropped
YEARS = [2025, 2027, 2030, 2035, 2040, 2045, 2050, 2055, 2060]


def raw_result(total_carbon, soil_carbon, unique_id='v1', recreation=None):
    variables = {
        'Treatment': ('', ['None', 'Thinning', 'None', 'None', 'FinalFelling', 'None', 'None', 'None', 'None']),
        TOTAL_CARBON: ('ton C/ha', total_carbon),
        SOIL_CARBON: ('ton C/ha', soil_carbon),
        'Not used in monetization': ('m', [1] * len(YEARS)),
        'UserDefined8 : UniqueID': ('', [unique_id] * len(YEARS)),
    }
    if recreation is not None:
        variables['Recreation Index After'] = ('', recreation)
    return raw_heureka_rows(YEARS, variables)


SPRUCE_BAU = raw_result([100, 101, 110, 120, 40, 50, 60, 70, 80], [50] * 9, recreation=[1] * 9)
SPRUCE_PRES = raw_result([100, 101, 115, 130, 145, 160, 175, 190, 205], [50, 50, 51, 52, 53, 54, 55, 56, 57],
                         recreation=[2] * 9)
PINE_BAU = raw_result([200, 201, 210, 220, 230, 240, 250, 260, 270], [60] * 9, recreation=[1] * 9)
PINE_PRES = raw_result([200, 201, 220, 240, 260, 280, 300, 320, 340], [60] * 9, recreation=[3] * 9)

SHEETS = {'P1 Spruce BAU': SPRUCE_BAU, 'P1 Spruce PRES': SPRUCE_PRES,
          'P1 Pine BAU': PINE_BAU, 'P1 Pine PRES': PINE_PRES}


def test_read_raw_heureka_results(heureka_results_factory):
    f = heureka_results_factory({'S': SPRUCE_BAU})
    result = read_raw_heureka_results(f, 'S')
    assert list(result['Year']) == [2025, 2030, 2035, 2040, 2045, 2050, 2055, 2060]
    assert list(result[TOTAL_CARBON])[:3] == [100, 110, 120]
    assert result.attrs[TOTAL_CARBON] == 'ton C/ha'
    assert 'Not used in monetization' in result.keys()


def test_read_raw_heureka_results_selected_variables(heureka_results_factory):
    f = heureka_results_factory({'S': SPRUCE_BAU})
    result = read_raw_heureka_results(f, 'S', read_only_these_variables=['Year', SOIL_CARBON])
    assert set(result.keys()) == {'Year', SOIL_CARBON}


def test_read_raw_heureka_results_without_year(heureka_results_factory):
    rows = [row for row in SPRUCE_BAU if row[1] != 'Year']
    f = heureka_results_factory({'S': rows})
    assert read_raw_heureka_results(f, 'S') is None


def test_read_raw_heureka_results_with_empty_rows(heureka_results_factory):
    rows = SPRUCE_BAU[:3] + [[None] * len(SPRUCE_BAU[0])] + SPRUCE_BAU[3:]
    f = heureka_results_factory({'S': rows})
    result = read_raw_heureka_results(f, 'S')
    assert SOIL_CARBON in result.keys()


def test_combine_raw_heureka_results(heureka_results_factory):
    f = heureka_results_factory({'S': SPRUCE_BAU})
    result = combine_raw_heureka_results(f, 'S', lambda x, y: x - y, {'x': TOTAL_CARBON, 'y': SOIL_CARBON})
    assert list(result)[:2] == [50, 60]


def test_combine_tables_copies_non_numeric_columns():
    t1 = pd.DataFrame({'Year': [2025, 2030], 'Treatment': ['None', 'X'], 'UserDefined8 : UniqueID': ['a', 'a'],
                       'C': [10., 20.]})
    t2 = pd.DataFrame({'Year': [2025, 2030], 'Treatment': ['None', 'Y'], 'UserDefined8 : UniqueID': ['a', 'a'],
                       'C': ['30', 40.]})  # numbers stored as text, as in output from the Woods platform
    combined = combine_tables([t1, t2], [0.25, 0.75])
    assert list(combined['Year']) == [2025, 2030]
    assert list(combined['Treatment']) == ['None', 'X']
    assert list(combined['UserDefined8 : UniqueID']) == ['a', 'a']
    assert list(combined['C']) == pytest.approx([25., 35.])


def test_rearrange_and_read_back(heureka_results_factory):
    f = heureka_results_factory(SHEETS)
    combine_sheets = {
        'P1 Combined Stands Spruce-and-Pine BAU': ['P1 Spruce BAU', 0.25, 'P1 Pine BAU', 0.75],
        'P1 Combined Stands Spruce-and-Pine PRES': ['P1 Spruce PRES', 0.25, 'P1 Pine PRES', 0.75]}
    wrote_monetization = rearrange_raw_heureka_results(f, list(SHEETS.keys()), combine_sheets)
    assert wrote_monetization is False

    assert 'Rearranged results' in openpyxl.load_workbook(f).sheetnames
    result = read_rearranged_heureka_results(f)
    assert list(result.keys()) == list(SHEETS.keys()) + list(combine_sheets.keys())

    spruce_bau = result['P1 Spruce BAU']
    assert TOTAL_CARBON in spruce_bau.keys()
    assert 'Not used in monetization' not in spruce_bau.keys()
    assert list(spruce_bau[TOTAL_CARBON])[:3] == [100, 110, 120]

    combined = result['P1 Combined Stands Spruce-and-Pine BAU']
    assert [float(x) for x in combined[TOTAL_CARBON]][:2] == pytest.approx([0.25 * 100 + 0.75 * 200,
                                                                             0.25 * 110 + 0.75 * 210])
    assert list(combined['UserDefined8 : UniqueID'])[0] == 'v1'
    assert list(combined['Year'])[0] == 2025


def test_rearrange_refuses_to_overwrite(heureka_results_factory, capsys):
    f = heureka_results_factory(SHEETS)
    rearrange_raw_heureka_results(f, ['P1 Spruce BAU'], None)
    before = read_rearranged_heureka_results(f)
    rearrange_raw_heureka_results(f, ['P1 Pine BAU'], None)
    assert 'already exists' in capsys.readouterr().out
    after = read_rearranged_heureka_results(f)
    assert list(after.keys()) == list(before.keys()) == ['P1 Spruce BAU']


def test_rearrange_missing_sheet_raises(heureka_results_factory):
    rows = [row for row in SPRUCE_BAU if row[1] != 'Year']
    f = heureka_results_factory({'No year': rows})
    with pytest.raises(ValueError):
        rearrange_raw_heureka_results(f, ['No year'], None)


def test_get_nature_and_climate_effect(heureka_results_factory):
    f = heureka_results_factory(SHEETS)
    carbon, dead_wood, recreation = get_nature_and_climate_effect(f, 'P1 Spruce', average_years=15)
    # Three 5 year periods: difference in (total - soil) between PRES and BAU
    # PRES: 50, 64, 78;  BAU: 50, 60, 70
    assert carbon == pytest.approx(0 + 4 + 8)
    assert dead_wood is None  # dead wood variables are not in the results
    assert recreation == pytest.approx(3.)


def test_get_nature_and_climate_effect_missing_sheet(heureka_results_factory):
    f = heureka_results_factory({'P1 Spruce BAU': SPRUCE_BAU})
    assert get_nature_and_climate_effect(f, 'P1 Spruce') == (None, None, None)
