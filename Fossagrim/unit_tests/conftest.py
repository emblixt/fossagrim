"""
Shared pytest fixtures.

Synthetic data
--------------
Small, made up, versions of the files used in the workflow, written to a temporary folder:
    forvaltningsplan_file:   a Forvaltningsplan with a 'data' sheet (header in row 8, as in the real files)
    heureka_results_factory: a "Heureka results.xlsx" file with raw results pasted into named sheets

Real data
---------
real_forvaltningsplan: the real FHF23-0007 Forvaltningsplan in Fossagrim/test_data. Tests using it are skipped
when the file is not present.
"""
import os

import matplotlib
import openpyxl
import pytest

matplotlib.use('Agg')  # never open plot windows during tests

TEST_DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'test_data')
REAL_FORVALTNINGSPLAN = os.path.join(TEST_DATA_DIR, 'FHF23-0007 Forvaltningsplan_v01_1.xlsx')
UNIT_TEST_DIR = os.path.dirname(os.path.abspath(__file__))

# Column names as they appear in the real files, including line breaks and trailing white space that
# read_excel() must strip
STAND_COLUMNS = [
    'HovedNr', 'Gnr', 'Bnr', ' Teig', 'Bestand', 'Miljøfig', 'Bonitering\ntreslag', 'Markslag', 'H.kl', 'Tetthet',
    'År', 'Prod.areal', 'Gran\n', 'Furu', 'Lauv', 'Total', 'Vol\n/\ndaa\n', 'Tilvekst\n/\ndaa', 'Areal status',
    'Volum status', 'Fossagrim ID', 'InventoryYear', 'CountyCode', 'Altitude', 'Latitude', 'SoilMoistureCode',
    'VegetationType', 'Peat', 'Antall trær pr mål', 'Gjennomsnitts diameter', 'Middelhøyde', 'Grunnflatesum',
    'Tretype_Alder_høyde', 'Svensk bonitet', 'Plantetetthet', 'Rotasjonsperiode', 'Tynnings år']


def stand_row(bestand, fossagrim_id, species, volum_status, prod_areal, gran, furu, lauv, alder=60.,
              vol_per_daa=20., markslag=13, soil_moisture=2, **overrides):
    """One row of the synthetic Forvaltningsplan, as a dict of column name: value"""
    row = {
        'HovedNr': 1, 'Gnr': 10, 'Bnr': 2, ' Teig': 1, 'Bestand': bestand, 'Miljøfig': 'Nei',
        'Bonitering\ntreslag': species, 'Markslag': markslag, 'H.kl': 5, 'Tetthet': 'Middels', 'År': alder,
        'Prod.areal': prod_areal, 'Gran\n': gran, 'Furu': furu, 'Lauv': lauv, 'Total': gran + furu + lauv,
        'Vol\n/\ndaa\n': vol_per_daa, 'Tilvekst\n/\ndaa': 0.5, 'Areal status': 1, 'Volum status': volum_status,
        'Fossagrim ID': fossagrim_id, 'InventoryYear': 2023, 'CountyCode': 42, 'Altitude': 300., 'Latitude': 59.5,
        'SoilMoistureCode': soil_moisture, 'VegetationType': 4, 'Peat': 0, 'Antall trær pr mål': 80.,
        'Gjennomsnitts diameter': 25., 'Middelhøyde': 18., 'Grunnflatesum': 30., 'Tretype_Alder_høyde': 'x',
        'Svensk bonitet': 20., 'Plantetetthet': 200., 'Rotasjonsperiode': 70., 'Tynnings år': 35.}
    row.update(overrides)
    return row


DEFAULT_STANDS = [
    # Two active spruce stands, one active pine stand, one inactive spruce stand and one unproductive stand
    stand_row(1, '1-1', 'Gran', 1, prod_areal=10., gran=180., furu=20., lauv=0., alder=60., vol_per_daa=20.),
    stand_row(2, '2-1', 'Gran', 1, prod_areal=30., gran=500., furu=50., lauv=50., alder=80., vol_per_daa=20.,
              markslag=14, soil_moisture=3),
    stand_row(3, '3-1', 'Furu', 1, prod_areal=20., gran=0., furu=300., lauv=0., alder=100., vol_per_daa=15.),
    stand_row(4, '4-1', 'Gran', 0, prod_areal=5., gran=50., furu=0., lauv=0.),
    stand_row(5, '5-1', 'Uproduktiv', 0, prod_areal=0., gran=0., furu=0., lauv=0.),
]


def write_forvaltningsplan(filename, stands, header_row=7, sheet_name='data'):
    """Writes a synthetic Forvaltningsplan. header_row is zero-based, as used by read_excel"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    for r in range(1, header_row + 1):
        ws.cell(r, 1).value = 'Some text above the table {}'.format(r)
    for c, key in enumerate(STAND_COLUMNS):
        ws.cell(header_row + 1, c + 1).value = key
    for r, stand in enumerate(stands):
        for c, key in enumerate(STAND_COLUMNS):
            ws.cell(header_row + 2 + r, c + 1).value = stand[key]
    wb.save(filename)
    return filename


@pytest.fixture
def forvaltningsplan_file(tmp_path):
    return write_forvaltningsplan(str(tmp_path / 'Test Forvaltningsplan.xlsx'), DEFAULT_STANDS)


def raw_heureka_rows(years, variables):
    """
    Rows of a raw Heureka result as pasted into Excel.

    :param years:
        list of the years of each period
    :param variables:
        dict of variable name: (unit, list of values, one per period)
    """
    rows = [['Variable', 'Category', 'Unit'] + ['Period {}'.format(i) for i in range(len(years))],
            ['Year', 'Treatments', 'year'] + list(years)]
    for name, (unit, values) in variables.items():
        rows.append([name, 'Results', unit] + list(values))
    return rows


def write_heureka_results(filename, sheets):
    """
    :param sheets:
        dict of sheet name: rows (see raw_heureka_rows)
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet_name, rows in sheets.items():
        ws = wb.create_sheet(sheet_name)
        for row in rows:
            ws.append(row)
    wb.save(filename)
    return filename


@pytest.fixture
def heureka_results_factory(tmp_path):
    def factory(sheets, name='Test Heureka results.xlsx'):
        return write_heureka_results(str(tmp_path / name), sheets)
    return factory


@pytest.fixture
def real_forvaltningsplan():
    if not os.path.isfile(REAL_FORVALTNINGSPLAN):
        pytest.skip('Real test data {} is not available'.format(os.path.basename(REAL_FORVALTNINGSPLAN)))
    return REAL_FORVALTNINGSPLAN
